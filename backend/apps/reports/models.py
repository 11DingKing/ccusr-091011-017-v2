"""
报表模型
"""
from django.conf import settings
from django.db import models


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
    分块导出任务：绑定创建时的数据快照，后续所有分块下载只能继续该快照。

    清单（file_specs / total_*）在创建时一次性写入且不再变动；
    即便之后权限或数据发生变化，已生成的清单也不会被悄悄改写。
    """

    STATUS_ACTIVE = 'active'
    STATUS_EXPIRED = 'expired'
    STATUS_CHOICES = [
        (STATUS_ACTIVE, '可下载'),
        (STATUS_EXPIRED, '已过期'),
    ]

    dataset = models.CharField('数据集', max_length=32)
    start_date = models.DateField('开始日期')
    end_date = models.DateField('结束日期')
    # 快照锚点：SQLite 序列号（单连接事务内即提交可见的最新事务）
    snapshot_seq = models.BigIntegerField('快照序列号')
    snapshot_taken_at = models.DateTimeField('快照创建时间')
    status = models.CharField('状态', max_length=16, choices=STATUS_CHOICES, default=STATUS_ACTIVE)
    file_specs = models.JSONField('文件清单', default=list)
    total_chunks = models.IntegerField('分块总数', default=0)
    total_rows = models.IntegerField('记录总条数', default=0)
    total_bytes = models.BigIntegerField('内容总字节数', default=0)
    chunk_size = models.IntegerField('分块行数', default=0)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True,
        related_name='export_jobs', verbose_name='创建人'
    )
    created_by_role = models.CharField('创建时角色', max_length=20, default='')
    expires_at = models.DateTimeField('过期时间')
    created_at = models.DateTimeField('创建时间', auto_now_add=True)

    class Meta:
        db_table = 'rp_export_job'
        verbose_name = '导出任务'
        verbose_name_plural = verbose_name
        ordering = ['-created_at']

    def __str__(self):
        return f"导出任务 {self.dataset} #{self.pk}"

    def is_expired(self, now=None):
        """是否已过期（含状态校准）。"""
        from django.utils import timezone
        current = now or timezone.now()
        return self.status == self.STATUS_EXPIRED or current >= self.expires_at

    def mark_expired(self):
        if self.status != self.STATUS_EXPIRED:
            ExportJob.objects.filter(pk=self.pk).update(status=self.STATUS_EXPIRED)
            self.status = self.STATUS_EXPIRED


class ExportChunk(models.Model):
    """导出任务下单个文件的单个分块；content 为冻结的 XLSX 字节，checksum 为其 SHA-256。"""

    job = models.ForeignKey(
        ExportJob, on_delete=models.CASCADE,
        related_name='chunks', verbose_name='导出任务'
    )
    file_index = models.IntegerField('文件序号')
    chunk_index = models.IntegerField('分块序号')
    row_count = models.IntegerField('分块记录数')
    content = models.BinaryField('分块内容')
    checksum = models.CharField('SHA-256校验值', max_length=64)
    size = models.IntegerField('字节数')
    created_at = models.DateTimeField('创建时间', auto_now_add=True)

    class Meta:
        db_table = 'rp_export_chunk'
        verbose_name = '导出分块'
        verbose_name_plural = verbose_name
        unique_together = [('job', 'file_index', 'chunk_index')]
        ordering = ['file_index', 'chunk_index']

    def __str__(self):
        return f"导出分块 {self.job_id} {self.file_index}-{self.chunk_index}"
