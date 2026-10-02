"""
报表URL配置
"""
from django.urls import path
from .views import (
    DashboardView,
    DailyReportView,
    ExportView,
    SnapshotExportChunkView,
    SnapshotExportCreateView,
    SnapshotExportManifestView,
    SystemMonitorView,
)

urlpatterns = [
    path('dashboard/', DashboardView.as_view(), name='dashboard'),
    path('daily-report/', DailyReportView.as_view(), name='daily-report'),
    path('export/', ExportView.as_view(), name='export'),

    # 固定快照的分块导出
    path('export/snapshots/', SnapshotExportCreateView.as_view(), name='export-snapshot-create'),
    path('export/snapshots/manifest/', SnapshotExportManifestView.as_view(), name='export-snapshot-manifest'),
    path(
        'export/snapshots/files/<int:file_index>/chunks/<int:chunk_index>/',
        SnapshotExportChunkView.as_view(),
        name='export-snapshot-chunk',
    ),

    path('system-monitor/', SystemMonitorView.as_view(), name='system-monitor'),
]
