"""
固定快照的分块导出服务。

设计要点：
1. 创建导出分三阶段：① 在单一 SQLite 读事务内物化全部数据行并记录快照版本号
   （事务期间其他连接提交的新记录不会进入本快照）；② 事务外渲染 XLSX 分块；
   ③ 短事务一次性写入任务、清单与分块。各文件因此共享同一数据版本（snapshot_seq），
   总数互相对得上；配合 WAL，导出不阻塞新记录写入。
2. 分块内容（XLSX 字节）与 SHA-256 校验值在创建时冻结入库；重复下载同一分块
   始终返回相同的字节与校验值。
3. 游标为 HMAC 签名令牌，绑定具体任务、快照版本与创建人；游标只能继续其签发时的
   快照，任务过期后所有后续请求被明确拒绝（410）。
4. 清单（文件/分块/总数）创建后不可变；权限变化只影响能否“新建”导出，不会改写
   已生成清单，也不影响创建人继续下载既有快照。
"""
import base64
import hashlib
import hmac
import io
import json
from datetime import timedelta

from django.conf import settings
from django.db import connection, transaction
from django.utils import timezone
from openpyxl import Workbook

from apps.warehouse.models import Approval, StockIn, StockOut
from .models import ExportChunk, ExportJob

# 默认/上限：每个分块包含的数据行数（表头不计）
DEFAULT_CHUNK_SIZE = 5000
MAX_CHUNK_SIZE = 20000
# 日期范围上限（天），避免一次性快照过大
MAX_RANGE_DAYS = 366


class ExportValidationError(Exception):
    """请求参数错误（映射为 400）。"""


class CursorError(Exception):
    """游标校验失败。kind: invalid | not_found | expired。"""

    def __init__(self, kind, message):
        self.kind = kind
        super().__init__(message)


# ==================== 数据集定义 ====================

def _unit_name_of(goods):
    variety = getattr(goods, 'variety', None)
    category = getattr(variety, 'category', None)
    unit = getattr(category, 'unit', None)
    return unit.name if unit else ''


def _stock_in_dataset(start_date, end_date):
    headers = ['货物编码', '货物名称', '入库数量', '单位', '批次号', '供应商', '操作人', '入库时间']
    qs = (
        StockIn.objects
        .filter(stock_in_time__date__range=(start_date, end_date))
        .order_by('id')
        .select_related('goods__variety__category__unit', 'operator')
    )

    def row(r):
        return [
            r.goods.code,
            r.goods.name,
            float(r.quantity),
            _unit_name_of(r.goods),
            r.batch_no,
            r.supplier,
            r.operator.username if r.operator else '',
            r.stock_in_time.strftime('%Y-%m-%d %H:%M:%S'),
        ]

    return '入库记录', '入库记录.xlsx', headers, qs, row


def _stock_out_dataset(start_date, end_date):
    headers = ['货物编码', '货物名称', '出库数量', '单位', '领用人', '领用部门',
               '状态', '操作人', '出库时间']
    qs = (
        StockOut.objects
        .filter(created_at__date__range=(start_date, end_date))
        .order_by('id')
        .select_related('goods__variety__category__unit', 'operator')
    )

    def row(r):
        return [
            r.goods.code,
            r.goods.name,
            float(r.quantity),
            _unit_name_of(r.goods),
            r.receiver,
            r.receiver_dept,
            r.get_status_display(),
            r.operator.username if r.operator else '',
            r.stock_out_time.strftime('%Y-%m-%d %H:%M:%S') if r.stock_out_time else '',
        ]

    return '出库记录', '出库记录.xlsx', headers, qs, row


def _approval_dataset(start_date, end_date):
    headers = ['审批ID', '出库单ID', '货物编码', '货物名称', '出库数量', '单位',
               '领用人', '申请人', '审批人', '审批状态', '审批意见', '审批时间']
    qs = (
        Approval.objects
        .filter(created_at__date__range=(start_date, end_date))
        .order_by('id')
        .select_related(
            'stock_out__goods__variety__category__unit',
            'stock_out__operator', 'approver',
        )
    )

    def row(r):
        out = r.stock_out
        goods = out.goods if out else None
        return [
            r.id,
            out.id if out else '',
            goods.code if goods else '',
            goods.name if goods else '',
            float(out.quantity) if out else '',
            _unit_name_of(goods) if goods else '',
            out.receiver if out else '',
            out.operator.username if out and out.operator else '',
            r.approver.username if r.approver else '',
            r.get_status_display(),
            r.remark,
            r.created_at.strftime('%Y-%m-%d %H:%M:%S'),
        ]

    return '审批记录', '审批记录.xlsx', headers, qs, row


DATASETS = {
    'stock_in': _stock_in_dataset,
    'stock_out': _stock_out_dataset,
    'approval': _approval_dataset,
}
# “收发和审批”一键导出
DATASET_GROUP_ALL = ['stock_in', 'stock_out', 'approval']


# ==================== 参数解析 ====================

def parse_params(params):
    """解析并校验创建请求参数，返回 (dataset_keys, start, end, chunk_size)。"""
    dataset_type = params.get('type', 'all')
    if dataset_type == 'all':
        dataset_keys = list(DATASET_GROUP_ALL)
    elif dataset_type in DATASETS:
        dataset_keys = [dataset_type]
    else:
        raise ExportValidationError(
            f"无效的导出类型：{dataset_type}，可选值为 all、stock_in、stock_out、approval"
        )

    today = timezone.localdate()
    try:
        start_raw = params.get('start_date')
        end_raw = params.get('end_date')
        start_date = (today - timedelta(days=29)) if not start_raw else _parse_date(start_raw)
        end_date = today if not end_raw else _parse_date(end_raw)
    except ValueError:
        raise ExportValidationError('日期格式无效，应为 YYYY-MM-DD')

    if start_date > end_date:
        raise ExportValidationError('开始日期不能晚于结束日期')
    if (end_date - start_date).days > MAX_RANGE_DAYS:
        raise ExportValidationError(f'导出日期范围不能超过 {MAX_RANGE_DAYS} 天')

    try:
        chunk_size = int(params.get('chunk_size', DEFAULT_CHUNK_SIZE))
    except (TypeError, ValueError):
        raise ExportValidationError('chunk_size 必须为正整数')
    if chunk_size < 1:
        raise ExportValidationError('chunk_size 必须为正整数')
    chunk_size = min(chunk_size, MAX_CHUNK_SIZE)

    return dataset_keys, start_date, end_date, chunk_size


def _parse_date(value):
    from datetime import datetime
    return datetime.strptime(value, '%Y-%m-%d').date()


# ==================== 快照 ====================

def _read_snapshot_seq():
    """
    返回当前读快照的版本锚点。

    SQLite 的 PRAGMA data_version 在读事务期间冻结为快照建立时的数据库版本号；
    其他连接在此之后的提交不会改变它，因此可作为快照标识。非 SQLite 环境返回 0，
    快照一致性退化为由数据库自身的事务隔离级别保证。
    """
    if connection.vendor != 'sqlite':
        return 0
    with connection.cursor() as cur:
        cur.execute('PRAGMA data_version')
        return int(cur.fetchone()[0])


def _render_xlsx(sheet_title, headers, rows):
    """将表头与一批数据行渲染为独立的 XLSX 字节。"""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(headers)
    for data_row in rows:
        ws.append(data_row)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ==================== 创建导出（快照冻结） ====================

def _snapshot_ttl():
    minutes = getattr(settings, 'EXPORT_SNAPSHOT_TTL_MINUTES', 120)
    return timedelta(minutes=minutes)


def _secret_key():
    return settings.SECRET_KEY.encode('utf-8')


def create_export(user, dataset_keys, start_date, end_date, chunk_size):
    """
    固定快照的分块导出，分三个阶段：

    1. 读快照（短事务）：在同一 SQLite 读事务内物化全部数据行并记录快照版本号，
       期间其他连接的提交不会进入本快照；
    2. 渲染（事务外）：把冻结的数据行渲染为 XLSX 分块并计算校验值，不占数据库锁；
    3. 落库（短事务）：任务、清单、分块一次性写入。

    这样新记录在导出期间可持续写入（WAL 下读写互不阻塞），而各文件分块始终
    取自阶段 1 的同一数据版本，总数互相一致。
    """
    now = timezone.now()
    expires_at = now + _snapshot_ttl()

    # ---- 阶段 1：在单一读快照内物化数据 ----
    with transaction.atomic():
        frozen_files = []
        for file_index, key in enumerate(dataset_keys):
            sheet_title, filename, headers, queryset, build_row = DATASETS[key](start_date, end_date)
            data_rows = [build_row(record) for record in queryset]
            frozen_files.append({
                'key': key,
                'file_index': file_index,
                'name': sheet_title,
                'filename': filename,
                'headers': headers,
                'rows': data_rows,
            })
        snapshot_seq = _read_snapshot_seq()

    # ---- 阶段 2：事务外渲染分块（内容与校验值在此冻结） ----
    rendered_files = []
    for frozen in frozen_files:
        rendered_chunks = []
        rows = frozen['rows']
        chunk_index = 1
        # max(len(rows), 1)：空数据集也生成一个仅含表头的分块
        for offset in range(0, max(len(rows), 1), chunk_size):
            part = rows[offset:offset + chunk_size]
            content = _render_xlsx(frozen['name'], frozen['headers'], part)
            rendered_chunks.append({
                'chunk_index': chunk_index,
                'row_count': len(part),
                'content': content,
                'checksum': hashlib.sha256(content).hexdigest(),
                'size': len(content),
            })
            chunk_index += 1

        rendered_files.append({**frozen, 'rendered_chunks': rendered_chunks})

    # ---- 阶段 3：短事务写入冻结结果 ----
    with transaction.atomic():
        job = ExportJob.objects.create(
            dataset='all' if len(dataset_keys) == len(DATASET_GROUP_ALL) else dataset_keys[0],
            start_date=start_date,
            end_date=end_date,
            snapshot_seq=snapshot_seq,
            snapshot_taken_at=now,
            status=ExportJob.STATUS_ACTIVE,
            file_specs=[],
            chunk_size=chunk_size,
            created_by=user,
            created_by_role=getattr(user, 'role', '') or '',
            expires_at=expires_at,
        )

        file_specs = []
        chunks_to_create = []
        total_rows = 0
        total_bytes = 0

        for rendered in rendered_files:
            file_bytes = 0
            for item in rendered['rendered_chunks']:
                chunks_to_create.append(ExportChunk(
                    job=job,
                    file_index=rendered['file_index'],
                    chunk_index=item['chunk_index'],
                    row_count=item['row_count'],
                    content=item['content'],
                    checksum=item['checksum'],
                    size=item['size'],
                ))
                file_bytes += item['size']
                total_rows += item['row_count']

            total_bytes += file_bytes
            file_specs.append({
                'file_index': rendered['file_index'],
                'dataset': rendered['key'],
                'name': rendered['name'],
                'filename': rendered['filename'],
                'rows': len(rendered['rows']),
                'chunks': len(rendered['rendered_chunks']),
                'size': file_bytes,
            })

        ExportChunk.objects.bulk_create(chunks_to_create)

        job.file_specs = file_specs
        job.total_chunks = len(chunks_to_create)
        job.total_rows = total_rows
        job.total_bytes = total_bytes
        job.save(update_fields=['file_specs', 'total_chunks', 'total_rows', 'total_bytes'])

        cursor = issue_cursor(job, user.id)

    return job, cursor


# ==================== 游标签发与解析 ====================

def issue_cursor(job, user_id):
    """签发绑定任务、快照与创建人的签名游标。"""
    payload = {
        'jid': job.id,
        'uid': user_id,
        'seq': job.snapshot_seq,
        'exp': int(job.expires_at.timestamp()),
    }
    raw = json.dumps(payload, separators=(',', ':'), sort_keys=True).encode('utf-8')
    body = base64.urlsafe_b64encode(raw).rstrip(b'=').decode('ascii')
    signature = hmac.new(_secret_key(), body.encode('ascii'), hashlib.sha256).hexdigest()
    return f'{body}.{signature}'


def resolve_cursor(cursor, user):
    """
    校验游标并返回对应任务；失败抛出 CursorError：
    - invalid: 游标缺失/格式错误/签名或内容被篡改/与创建人不匹配
    - not_found: 任务不存在
    - expired: 任务（或游标自身过期时间）已过期

    游标只在创建人本人会话内有效，避免游标泄露后被他人重放。
    """
    if not cursor or not isinstance(cursor, str) or '.' not in cursor:
        raise CursorError('invalid', '无效的导出游标')

    body, _, signature = cursor.rpartition('.')
    expected = hmac.new(_secret_key(), body.encode('ascii'), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(signature, expected):
        raise CursorError('invalid', '导出游标签名无效')

    try:
        padded = body + '=' * (-len(body) % 4)
        payload = json.loads(base64.urlsafe_b64decode(padded.encode('ascii')))
        job_id = int(payload['jid'])
        user_id = int(payload['uid'])
        snapshot_seq = int(payload['seq'])
        exp = int(payload['exp'])
    except (ValueError, TypeError, KeyError, json.JSONDecodeError):
        raise CursorError('invalid', '导出游标内容无效')

    try:
        job = ExportJob.objects.get(pk=job_id)
    except ExportJob.DoesNotExist:
        raise CursorError('not_found', '导出任务不存在或已被清理')

    # 游标必须由创建人本人继续使用
    if job.created_by_id != user_id or not getattr(user, 'id', None) or user.id != user_id:
        raise CursorError('invalid', '游标不属于当前用户，只能继续本人创建的导出快照')

    # 游标必须继续其签发时的快照：快照锚点不一致即拒绝
    if job.snapshot_seq != snapshot_seq:
        raise CursorError('invalid', '游标与导出任务快照不匹配')

    now = timezone.now()
    if now.timestamp() >= exp or job.is_expired(now):
        job.mark_expired()
        raise CursorError(
            'expired',
            f'导出快照已过期（有效截止 {timezone.localtime(job.expires_at).strftime("%Y-%m-%d %H:%M:%S")}），请重新发起导出'
        )

    return job


# ==================== 清单与分块读取 ====================

def build_manifest(job, cursor):
    """从冻结数据构建清单摘要；其内容在任务创建后恒定。"""
    chunk_entries = list(
        ExportChunk.objects.filter(job=job)
        .order_by('file_index', 'chunk_index')
        .values('file_index', 'chunk_index', 'row_count', 'size', 'checksum')
    )
    return {
        'export_id': job.id,
        'cursor': cursor,
        'snapshot': {
            'seq': job.snapshot_seq,
            'taken_at': timezone.localtime(job.snapshot_taken_at).strftime('%Y-%m-%d %H:%M:%S'),
            'expires_at': timezone.localtime(job.expires_at).strftime('%Y-%m-%d %H:%M:%S'),
            'expired': job.is_expired(),
        },
        'dataset': job.dataset,
        'start_date': job.start_date.strftime('%Y-%m-%d'),
        'end_date': job.end_date.strftime('%Y-%m-%d'),
        'chunk_size': job.chunk_size,
        'total_files': len(job.file_specs),
        'total_chunks': job.total_chunks,
        'total_rows': job.total_rows,
        'total_bytes': job.total_bytes,
        'files': job.file_specs,
        'chunks': [
            {
                'file_index': c['file_index'],
                'chunk_index': c['chunk_index'],
                'rows': c['row_count'],
                'size': c['size'],
                'checksum': c['checksum'],
            }
            for c in chunk_entries
        ],
    }


def get_chunk(job, file_index, chunk_index):
    """取任务内指定分块（冻结字节）；不存在返回 None。"""
    return (
        ExportChunk.objects
        .filter(job=job, file_index=file_index, chunk_index=chunk_index)
        .first()
    )
