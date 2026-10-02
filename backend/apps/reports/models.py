"""
报表模型
"""
from django.conf import settings
from django.db import models
from django.utils import timezone


class DailyReport(models.Model):
    """每日报表模型"""
    report_date = models.DateField('报表日期', unique=True)
    in_count = models.IntegerField('入库次数', default=0)
    in_total = models.DecimalField('入库总量', max_digits=12, decimal_places=2, default=0)
    out_count = models.IntegerField('出库次数', default=0)
    out_total = models.DecimalField('出库总量', max_digits=12, decimal_places=2, default=0)
    warning_count = models.IntegerField('预警数量', default=0)
    summary = models.TextField('报表摘要', blank=True)
    created_at = models.DateTimeField('创建时间', auto_now_add=True)

    class Meta:
        db_table = 'rp_daily_report'
        verbose_name = '每日报表'
        verbose_name_plural = verbose_name
        ordering = ['-report_date']

    def __str__(self):
        return f"报表 - {self.report_date}"


class ExportJob(models.Model):
    """
    快照导出任务。

    创建时把本次导出要覆盖的数据范围（类型、日期区间、结果行的完整内容与顺序）
    冻结到 ExportSnapshotRow；后续所有分块只读取冻结内容，不受新增记录、
    审批状态变更或权限调整影响。
    """
    STATUS_PENDING = 'pending'
    STATUS_COMPLETED = 'completed'
    STATUS_EXPIRED = 'expired'
    STATUS_CHOICES = [
        (STATUS_PENDING, '待下载'),
        (STATUS_COMPLETED, '已完成'),
        (STATUS_EXPIRED, '已过期'),
    ]

    creator = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True,
        related_name='export_jobs', verbose_name='创建人'
    )
    export_type = models.CharField('导出类型', max_length=100)
    start_date = models.DateField('开始日期')
    end_date = models.DateField('结束日期')
    # 每个分区（入库/出库/审批）在快照中的行数，创建后永不改变
    section_counts = models.JSONField('分区行数', default=dict)
    total_rows = models.IntegerField('总行数', default=0)
    chunk_size = models.IntegerField('每块行数', default=1000)
    status = models.CharField('状态', max_length=20, choices=STATUS_CHOICES, default=STATUS_PENDING)
    cursor_token = models.CharField('游标令牌', max_length=64, unique=True)
    created_at = models.DateTimeField('快照创建时间', auto_now_add=True)
    expires_at = models.DateTimeField('过期时间')
    completed_at = models.DateTimeField('完成时间', null=True, blank=True)

    class Meta:
        db_table = 'rp_export_job'
        verbose_name = '快照导出任务'
        verbose_name_plural = verbose_name
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['cursor_token']),
            models.Index(fields=['status', 'expires_at']),
        ]

    def __str__(self):
        return f"导出任务 {self.export_type} - {self.id}"

    @property
    def total_chunks(self):
        if self.total_rows == 0:
            return 1
        return (self.total_rows + self.chunk_size - 1) // self.chunk_size

    def is_expired(self, now=None):
        if now is None:
            now = timezone.now()
        return self.expires_at <= now

    def chunk_indexes(self):
        """返回 (块序号, 起始下标, 结束下标) 列表，下标基于冻结行的 position。"""
        result = []
        index = 0
        chunk_no = 0
        while index < self.total_rows:
            result.append((chunk_no, index, min(index + self.chunk_size, self.total_rows)))
            index += self.chunk_size
            chunk_no += 1
        if not result:
            result.append((0, 0, 0))
        return result


class ExportSnapshotRow(models.Model):
    """创建任务时冻结的一行数据，内容与顺序在快照生命周期内永不改变。"""
    job = models.ForeignKey(
        ExportJob, on_delete=models.CASCADE,
        related_name='rows', verbose_name='导出任务'
    )
    position = models.IntegerField('行序号')
    payload = models.JSONField('冻结行内容')

    class Meta:
        db_table = 'rp_export_snapshot_row'
        verbose_name = '导出快照行'
        verbose_name_plural = verbose_name
        unique_together = [('job', 'position')]
        ordering = ['position']
        indexes = [models.Index(fields=['job', 'position'])]

    def __str__(self):
        return f"快照行 {self.job_id}-{self.position}"


class ExportChunk(models.Model):
    """导出分块：首次下载时按冻结行渲染并固化，重复下载返回相同字节与校验值。"""
    job = models.ForeignKey(
        ExportJob, on_delete=models.CASCADE,
        related_name='chunks', verbose_name='导出任务'
    )
    chunk_index = models.IntegerField('分块序号')
    start_index = models.IntegerField('起始下标')
    end_index = models.IntegerField('结束下标')
    row_count = models.IntegerField('行数', default=0)
    file_name = models.CharField('文件名', max_length=200)
    content = models.BinaryField('文件内容')
    checksum = models.CharField('SHA256校验值', max_length=64)
    size_bytes = models.IntegerField('字节数', default=0)
    created_at = models.DateTimeField('创建时间', auto_now_add=True)

    class Meta:
        db_table = 'rp_export_chunk'
        verbose_name = '导出分块'
        verbose_name_plural = verbose_name
        unique_together = [('job', 'chunk_index')]
        ordering = ['chunk_index']

    def __str__(self):
        return f"导出分块 {self.job_id}-{self.chunk_index}"
