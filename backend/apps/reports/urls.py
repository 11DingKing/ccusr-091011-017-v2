"""
报表URL配置
"""
from django.urls import path
from .views import (
    DashboardView, DailyReportView, ExportView, SystemMonitorView,
    SnapshotExportCreateView, SnapshotExportStatusView, SnapshotExportChunkView,
)

urlpatterns = [
    path('dashboard/', DashboardView.as_view(), name='dashboard'),
    path('daily-report/', DailyReportView.as_view(), name='daily-report'),
    path('export/', ExportView.as_view(), name='export'),
    path('system-monitor/', SystemMonitorView.as_view(), name='system-monitor'),

    # 固定快照分块导出
    path('exports/', SnapshotExportCreateView.as_view(), name='snapshot-export-create'),
    path('exports/status/', SnapshotExportStatusView.as_view(), name='snapshot-export-status'),
    path('exports/chunks/', SnapshotExportChunkView.as_view(), name='snapshot-export-chunk'),
]
