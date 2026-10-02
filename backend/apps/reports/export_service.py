"""
固定快照分块导出服务。

流程：
1. 创建任务时在单个事务内读取数据并把每一行完整冻结到 ExportSnapshotRow，
   同时固化分区行数、分块方案、游标令牌与过期时间，返回清单摘要。
2. 后续请求只能凭游标访问该快照；分块首次下载时按冻结行渲染 Excel 并把
   字节与 SHA256 校验值落库，重复下载原样返回，保证内容与校验值恒定。
3. 快照过期后一律拒绝（410），不会续期，也不会回源读取业务表。
"""
import hashlib
import io
import secrets
from datetime import datetime, timedelta

from django.conf import settings
from django.db import IntegrityError, transaction
from django.utils import timezone, dateformat
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from apps.core.exceptions import (
    BusinessException, NotFoundException, PermissionException, SnapshotExpiredException,
)
from apps.warehouse.models import Approval, StockIn, StockOut
from .models import ExportChunk, ExportJob, ExportSnapshotRow

# 支持的导出分区
SECTION_STOCK_IN = 'stock_in'
SECTION_STOCK_OUT = 'stock_out'
SECTION_APPROVAL = 'approval'

EXPORT_TYPES = {
    SECTION_STOCK_IN: {
        'label': '入库记录',
        'sections': [SECTION_STOCK_IN],
    },
    SECTION_STOCK_OUT: {
        'label': '出库记录',
        'sections': [SECTION_STOCK_OUT],
    },
    SECTION_APPROVAL: {
        'label': '审批记录',
        'sections': [SECTION_APPROVAL],
    },
    'all': {
        'label': '收发和审批数据',
        'sections': [SECTION_STOCK_IN, SECTION_STOCK_OUT, SECTION_APPROVAL],
    },
}

SECTION_LABELS = {
    SECTION_STOCK_IN: '入库记录',
    SECTION_STOCK_OUT: '出库记录',
    SECTION_APPROVAL: '审批记录',
}

SECTION_HEADERS = {
    SECTION_STOCK_IN: [
        '记录ID', '货物编码', '货物名称', '入库数量', '单位', '批次号',
        '供应商', '操作人', '入库时间',
    ],
    SECTION_STOCK_OUT: [
        '记录ID', '货物编码', '货物名称', '出库数量', '单位', '领用人',
        '领用部门', '状态', '操作人', '出库时间',
    ],
    SECTION_APPROVAL: [
        '审批ID', '出库记录ID', '货物编码', '货物名称', '出库数量', '单位',
        '领用人', '领用部门', '审批状态', '审批人', '审批意见',
        '申请时间', '审批时间',
    ],
}

DEFAULT_CHUNK_SIZE = 1000
MAX_CHUNK_SIZE = 5000
MAX_RANGE_DAYS = 366
SNAPSHOT_TTL_SECONDS = getattr(settings, 'EXPORT_SNAPSHOT_TTL_SECONDS', 3600)

HEADER_FONT = Font(bold=True, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='0066FF', end_color='0066FF', fill_type='solid')
HEADER_ALIGNMENT = Alignment(horizontal='center', vertical='center')
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin'),
)
# 固定文档元数据时间，避免任何环境相关信息进入文件字节
FIXED_DOC_TIME = datetime(2000, 1, 1, 0, 0, 0)


class ExportValidationError(BusinessException):
    """导出参数错误（400）。"""
    def __init__(self, message):
        super().__init__(message, code=400)


def _decimal(value):
    """Decimal 不能直接放入 JSONField，统一转为 float。"""
    return float(value) if value is not None else None


def _dt(value):
    return value.strftime('%Y-%m-%d %H:%M:%S') if value else ''


def _unit_name(goods):
    variety = getattr(goods, 'variety', None)
    category = getattr(variety, 'category', None) if variety else None
    unit = getattr(category, 'unit', None) if category else None
    return unit.name if unit else ''


def _stock_in_rows(start, end):
    queryset = StockIn.objects.select_related(
        'goods', 'goods__variety__category__unit', 'operator'
    ).filter(
        stock_in_time__date__gte=start,
        stock_in_time__date__lte=end,
    ).order_by('stock_in_time', 'id')
    for r in queryset.iterator():
        goods = r.goods
        yield [
            r.id,
            goods.code if goods else '',
            goods.name if goods else '',
            _decimal(r.quantity),
            _unit_name(goods),
            r.batch_no or '',
            r.supplier or '',
            r.operator.username if r.operator else '',
            _dt(r.stock_in_time),
        ]


def _stock_out_rows(start, end):
    queryset = StockOut.objects.select_related(
        'goods', 'goods__variety__category__unit', 'operator'
    ).filter(
        created_at__date__gte=start,
        created_at__date__lte=end,
    ).order_by('created_at', 'id')
    for r in queryset.iterator():
        goods = r.goods
        yield [
            r.id,
            goods.code if goods else '',
            goods.name if goods else '',
            _decimal(r.quantity),
            _unit_name(goods),
            r.receiver or '',
            r.receiver_dept or '',
            r.get_status_display(),
            r.operator.username if r.operator else '',
            _dt(r.stock_out_time),
        ]


def _approval_rows(start, end):
    queryset = Approval.objects.select_related(
        'stock_out', 'stock_out__goods',
        'stock_out__goods__variety__category__unit',
        'stock_out__operator', 'approver',
    ).filter(
        created_at__date__gte=start,
        created_at__date__lte=end,
    ).order_by('created_at', 'id')
    for a in queryset.iterator():
        out = a.stock_out
        goods = out.goods if out else None
        yield [
            a.id,
            out.id if out else '',
            goods.code if goods else '',
            goods.name if goods else '',
            _decimal(out.quantity) if out else None,
            _unit_name(goods) if goods else '',
            out.receiver if out else '',
            out.receiver_dept if out else '',
            a.get_status_display(),
            a.approver.username if a.approver else '',
            a.remark or '',
            _dt(out.created_at) if out else '',
            _dt(a.updated_at),
        ]


_SECTION_BUILDERS = {
    SECTION_STOCK_IN: _stock_in_rows,
    SECTION_STOCK_OUT: _stock_out_rows,
    SECTION_APPROVAL: _approval_rows,
}


def parse_date(params, name, default):
    raw = params.get(name)
    if not raw:
        return default
    try:
        return datetime.strptime(raw, '%Y-%m-%d').date()
    except (TypeError, ValueError):
        raise ExportValidationError(f'{name} 格式应为 YYYY-MM-DD')


def parse_chunk_size(params):
    raw = params.get('chunk_size')
    if not raw:
        return DEFAULT_CHUNK_SIZE
    try:
        value = int(raw)
    except (TypeError, ValueError):
        raise ExportValidationError('chunk_size 必须为正整数')
    if value < 1:
        raise ExportValidationError('chunk_size 必须为正整数')
    return min(value, MAX_CHUNK_SIZE)


@transaction.atomic
def create_export_job(user, export_type, params):
    """
    创建固定快照导出任务。

    权限在创建时按当前用户校验；通过后数据被冻结，之后该用户权限变化
    不会影响本任务，也不会改写已冻结的清单。
    """
    config = EXPORT_TYPES.get(export_type)
    if config is None:
        raise ExportValidationError('无效的导出类型')

    # 权限仅在创建任务时按当前状态校验一次，随后冻结到任务（creator）。
    # 用户之后的角色变化只影响其新建导出，不会改动本任务的清单与可访问性。
    if SECTION_APPROVAL in config['sections'] and not user.is_admin:
        raise PermissionException('审批数据仅管理员可导出')

    today = timezone.localdate()
    start = parse_date(params, 'start_date', today - timedelta(days=30))
    end = parse_date(params, 'end_date', today)
    if end < start:
        raise ExportValidationError('end_date 不能早于 start_date')
    if (end - start).days > MAX_RANGE_DAYS:
        raise ExportValidationError(f'导出区间不能超过 {MAX_RANGE_DAYS} 天')

    chunk_size = parse_chunk_size(params)

    now = timezone.now()
    token = secrets.token_urlsafe(32)
    # section_counts/total_rows 在冻结过程中回填；整个过程处于同一事务，
    # 对外可见时清单已完整且不会变化。
    job = ExportJob.objects.create(
        creator=user if user and user.is_authenticated else None,
        export_type=export_type,
        start_date=start,
        end_date=end,
        section_counts={},
        total_rows=0,
        chunk_size=chunk_size,
        cursor_token=token,
        expires_at=now + timedelta(seconds=SNAPSHOT_TTL_SECONDS),
    )

    # 按分区顺序流式冻结，position 为全局连续下标；分批落库避免整月数据驻留内存
    section_counts = {}
    position = 0
    batch = []
    for section in config['sections']:
        section_total = 0
        for cells in _SECTION_BUILDERS[section](start, end):
            batch.append(ExportSnapshotRow(
                job=job, position=position,
                payload={'section': section, 'cells': cells},
            ))
            position += 1
            section_total += 1
            if len(batch) >= MAX_CHUNK_SIZE:
                ExportSnapshotRow.objects.bulk_create(batch, batch_size=MAX_CHUNK_SIZE)
                batch = []
        section_counts[section] = section_total
    if batch:
        ExportSnapshotRow.objects.bulk_create(batch, batch_size=MAX_CHUNK_SIZE)

    job.section_counts = section_counts
    job.total_rows = position
    job.save(update_fields=['section_counts', 'total_rows'])

    return job


def _chunk_section_counts(job, start_index, end_index):
    """统计一个分块 [start_index, end_index) 内各分区的行数。"""
    counts = {}
    cursor = 0
    # 分区及顺序以快照固化的 section_counts 为准，不依赖当前代码配置
    for section in job.section_counts.keys():
        total = cursor + job.section_counts.get(section, 0)
        counts[section] = max(0, min(end_index, total) - max(start_index, cursor))
        cursor = total
    return counts


def build_manifest(job):
    """首个请求返回的清单摘要，内容全部来自固化字段，不回源业务表。"""
    chunks = []
    for index, start, end in job.chunk_indexes():
        chunks.append({
            'index': index,
            'start_position': start,
            'end_position': end,
            'row_count': end - start,
            'sections': _chunk_section_counts(job, start, end),
            'download_url': f'/api/exports/chunks/?cursor={job.cursor_token}&chunk={index}',
        })
    return {
        'export_job_id': job.id,
        'cursor': job.cursor_token,
        'export_type': job.export_type,
        'export_type_label': EXPORT_TYPES.get(job.export_type, {}).get('label', job.export_type),
        'start_date': job.start_date.strftime('%Y-%m-%d'),
        'end_date': job.end_date.strftime('%Y-%m-%d'),
        'snapshot_created_at': dateformat.format(timezone.localtime(job.created_at), 'Y-m-d H:i:s'),
        'expires_at': dateformat.format(timezone.localtime(job.expires_at), 'Y-m-d H:i:s'),
        'expires_in_seconds': max(0, int((job.expires_at - timezone.now()).total_seconds())),
        'chunk_size': job.chunk_size,
        'total_chunks': job.total_chunks,
        'total_rows': job.total_rows,
        'sections': [
            {'key': key, 'label': SECTION_LABELS[key], 'rows': job.section_counts.get(key, 0)}
            for key in job.section_counts.keys()
        ],
        'chunks': chunks,
    }


def get_accessible_job(cursor, user):
    """
    凭游标取出可访问的任务。游标无效/非本人 → 404（不泄露任务是否存在）；
    已过期 → 410 明确拒绝。
    """
    if not cursor:
        raise ExportValidationError('缺少 cursor 游标参数')
    job = ExportJob.objects.filter(cursor_token=cursor).first()
    if job is None or user is None or job.creator_id != user.id:
        raise NotFoundException('导出任务不存在或无权访问')
    if job.is_expired() or job.status == ExportJob.STATUS_EXPIRED:
        if job.status != ExportJob.STATUS_EXPIRED:
            job.status = ExportJob.STATUS_EXPIRED
            job.save(update_fields=['status'])
        raise SnapshotExpiredException('导出快照已过期，请重新创建导出任务')
    return job


def get_chunk_index(job, params):
    raw = params.get('chunk', '0')
    try:
        index = int(raw)
    except (TypeError, ValueError):
        raise ExportValidationError('chunk 必须为非负整数')
    if index < 0 or index >= job.total_chunks:
        raise ExportValidationError(
            f'chunk 超出范围，本任务共 {job.total_chunks} 个分块（0 ~ {job.total_chunks - 1}）'
        )
    return index


def chunk_file_name(job, index):
    label = EXPORT_TYPES[job.export_type]['label']
    return (
        f"{label}_{job.start_date.strftime('%Y%m%d')}-{job.end_date.strftime('%Y%m%d')}"
        f"_{job.cursor_token[:8]}_part{index + 1:04d}of{job.total_chunks:04d}.xlsx"
    )


def _style_header(ws, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER


def _autosize(ws):
    for column in ws.columns:
        max_length = 0
        letter = column[0].column_letter
        for cell in column:
            value = '' if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        ws.column_dimensions[letter].width = min((max_length + 2) * 1.2, 60)


def render_chunk(job, chunk_index, start_index, end_index):
    """
    用冻结行渲染 Excel 字节。输入完全来自快照，输出确定：
    相同的冻结行 + 相同的块参数必然产生相同字节。
    """
    sections = list(job.section_counts.keys())
    rows = list(
        ExportSnapshotRow.objects.filter(
            job=job, position__gte=start_index, position__lt=end_index
        ).order_by('position').values_list('payload', flat=True)
    )

    wb = Workbook()
    wb.properties.creator = 'export-service'
    wb.properties.created = FIXED_DOC_TIME
    wb.properties.modified = FIXED_DOC_TIME

    first_sheet = True
    for section in sections:
        ws = wb.active if first_sheet else wb.create_sheet()
        first_sheet = False
        ws.title = SECTION_LABELS[section]
        ws.append(SECTION_HEADERS[section])
        _style_header(ws, len(SECTION_HEADERS[section]))
        for payload in rows:
            if payload['section'] != section:
                continue
            ws.append(payload['cells'])
        _autosize(ws)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


@transaction.atomic
def get_or_build_chunk(job, chunk_index):
    """
    取分块记录；首次下载时渲染并固化 content/checksum。
    (job, chunk_index) 唯一约束兜底并发：重复提交时必有一方渲染失败，
    转而回读另一方已固化的结果，保证对外字节与校验值唯一。
    """
    chunk = ExportChunk.objects.filter(job=job, chunk_index=chunk_index).first()
    if chunk is not None:
        return chunk, False

    starts = {idx: (start, end) for idx, start, end in job.chunk_indexes()}
    start_index, end_index = starts[chunk_index]
    content = render_chunk(job, chunk_index, start_index, end_index)
    checksum = hashlib.sha256(content).hexdigest()

    try:
        with transaction.atomic():
            chunk = ExportChunk.objects.create(
                job=job,
                chunk_index=chunk_index,
                start_index=start_index,
                end_index=end_index,
                row_count=end_index - start_index,
                file_name=chunk_file_name(job, chunk_index),
                content=content,
                checksum=checksum,
                size_bytes=len(content),
            )
    except IntegrityError:
        # 并发首次下载：另一请求已固化该分块，回读其结果
        chunk = ExportChunk.objects.get(job=job, chunk_index=chunk_index)
        return chunk, False

    built = ExportChunk.objects.filter(job=job).count()
    if built >= job.total_chunks and job.status == ExportJob.STATUS_PENDING:
        job.status = ExportJob.STATUS_COMPLETED
        job.completed_at = timezone.now()
        job.save(update_fields=['status', 'completed_at'])

    return chunk, True


def build_chunk_state(job):
    """清单中各分块当前下载状态（供状态查询）。"""
    built = {c.chunk_index: c for c in ExportChunk.objects.filter(job=job)}
    result = []
    for index, start, end in job.chunk_indexes():
        chunk = built.get(index)
        result.append({
            'index': index,
            'row_count': end - start,
            'sections': _chunk_section_counts(job, start, end),
            'ready': chunk is not None,
            'checksum': chunk.checksum if chunk else None,
            'size_bytes': chunk.size_bytes if chunk else None,
            'file_name': chunk.file_name if chunk else chunk_file_name(job, index),
        })
    return result
