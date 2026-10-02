"""
清理过期的导出快照任务及其冻结数据（含已渲染分块）。
"""
from django.core.management.base import BaseCommand
from django.utils import timezone

from apps.reports.models import ExportJob


class Command(BaseCommand):
    help = '清理过期的导出快照任务（ExportJob 级联删除快照行与分块文件）'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run', action='store_true',
            help='只统计将清理的任务数，不实际删除',
        )

    def handle(self, *args, **options):
        now = timezone.now()
        expired_qs = ExportJob.objects.filter(expires_at__lt=now)
        count = expired_qs.count()

        if options['dry_run']:
            self.stdout.write(f'预计清理过期导出快照任务: {count} 个')
            return

        if count:
            # CASCADE 会一并删除 ExportSnapshotRow 与 ExportChunk
            deleted = expired_qs.delete()[0]
        else:
            deleted = 0
        self.stdout.write(self.style.SUCCESS(f'已清理过期导出快照任务: {count} 个（删除对象 {deleted} 项）'))
