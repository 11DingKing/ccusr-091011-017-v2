"""固定快照分块导出的验收测试。"""
import hashlib
import io
from datetime import timedelta
from decimal import Decimal

from django.utils import timezone
from openpyxl import load_workbook
from rest_framework.test import APIClient

from apps.authentication.backends import generate_token
from apps.authentication.models import User
from apps.warehouse.models import Approval, Category, Goods, StockIn, StockOut, Unit, Variety
from .models import ExportJob


class SnapshotExportTestBase:
    def setUp(self):
        self.admin = User.objects.create_user('export-admin', 'testpass123', role='admin')
        self.client = APIClient()
        self.client.credentials(HTTP_AUTHORIZATION=f'Bearer {generate_token(self.admin)}')

        unit = Unit.objects.create(name='件', created_by=self.admin)
        category = Category.objects.create(name='监管器材', unit=unit, created_by=self.admin)
        variety = Variety.objects.create(name='封存设备', category=category, created_by=self.admin)
        self.goods = Goods.objects.create(
            variety=variety, name='封存终端', code='SNAP-001',
            quantity=Decimal('10'), warning_threshold=Decimal('5'),
        )

    def _range_params(self, **extra):
        today = timezone.localdate()
        params = {
            'start_date': (today - timedelta(days=1)).strftime('%Y-%m-%d'),
            'end_date': today.strftime('%Y-%m-%d'),
        }
        params.update(extra)
        return params

    def _add_records(self, suffix):
        """写入一条入库、一条出库及对应审批，返回各模型实例。"""
        stock_in = StockIn.objects.create(
            goods=self.goods, operator=self.admin,
            quantity=Decimal('3'), batch_no=f'B-{suffix}', supplier='某单位',
        )
        stock_out = StockOut.objects.create(
            goods=self.goods, operator=self.admin, receiver=f'领用人{suffix}',
            quantity=Decimal('1'), status='pending',
        )
        approval = Approval.objects.create(
            stock_out=stock_out, approver=self.admin, status='pending', remark='待审',
        )
        return stock_in, stock_out, approval

    def _create_export(self, **extra):
        response = self.client.post('/api/export/snapshots/', self._range_params(**extra), format='json')
        assert response.status_code == 200, response.content
        return response.json()['data']

    def _download(self, manifest, file_index, chunk_index):
        return self.client.get(
            f'/api/export/snapshots/files/{file_index}/chunks/{chunk_index}/',
            {'cursor': manifest['cursor']},
        )


from django.test import TransactionTestCase


class SnapshotExportTests(SnapshotExportTestBase, TransactionTestCase):
    def test_create_returns_manifest_summary_and_cursor(self):
        self._add_records('1')
        manifest = self._create_export()

        self.assertTrue(manifest['cursor'])
        self.assertEqual(manifest['total_files'], 3)
        self.assertEqual(manifest['dataset'], 'all')
        # 入库 / 出库 / 审批 各 1 条
        rows_per_file = {f['file_index']: f['rows'] for f in manifest['files']}
        self.assertEqual(rows_per_file, {0: 1, 1: 1, 2: 1})
        self.assertEqual(manifest['total_rows'], 3)
        self.assertEqual(manifest['total_chunks'], 3)
        # 清单内每文件分块校验值齐全
        for chunk in manifest['chunks']:
            self.assertEqual(len(chunk['checksum']), 64)
        self.assertFalse(manifest['snapshot']['expired'])

    def test_new_records_during_download_do_not_change_manifest_or_chunks(self):
        self._add_records('before')
        manifest = self._create_export(chunk_size=10)
        cursor = manifest['cursor']

        # 模拟“长时间分页下载”：边下载边有新记录进入
        first = self._download(manifest, 0, 1)
        self.assertEqual(first.status_code, 200)
        self._add_records('during-1')
        self._add_records('during-2')

        # 再次取清单：仍固定在原快照，总数不变
        again = self.client.get('/api/export/snapshots/manifest/', {'cursor': cursor})
        self.assertEqual(again.status_code, 200)
        again_data = again.json()['data']
        self.assertEqual(again_data['total_rows'], manifest['total_rows'])
        self.assertEqual(again_data['chunks'], manifest['chunks'])
        self.assertEqual(again_data['snapshot']['seq'], manifest['snapshot']['seq'])

        # 逐块下载，字节与校验值必须和清单一致
        downloaded = {}
        for spec in manifest['chunks']:
            resp = self._download(manifest, spec['file_index'], spec['chunk_index'])
            self.assertEqual(resp.status_code, 200)
            content = b''.join(resp.streaming_content) if resp.streaming else resp.content
            downloaded[(spec['file_index'], spec['chunk_index'])] = content
            self.assertEqual(hashlib.sha256(content).hexdigest(), spec['checksum'])
            self.assertEqual(resp['X-Checksum-Sha256'], spec['checksum'])
            self.assertEqual(resp['ETag'], f'"{spec["checksum"]}"')

        # 块内数据行数与清单声明一致（XLSX 首行表头不计）
        for spec in manifest['chunks']:
            wb = load_workbook(io.BytesIO(downloaded[(spec['file_index'], spec['chunk_index'])]))
            ws = wb.active
            self.assertEqual(ws.max_row - 1, spec['rows'])

        # 文件行数之和 == 清单总行数
        self.assertEqual(sum(f['rows'] for f in manifest['files']), manifest['total_rows'])

        # 新记录对“之后新建的导出”可见，但不在旧快照里
        new_manifest = self._create_export(chunk_size=10)
        self.assertEqual(new_manifest['total_rows'], 9)
        self.assertGreaterEqual(new_manifest['snapshot']['seq'], manifest['snapshot']['seq'])

    def test_repeated_chunk_download_is_byte_identical(self):
        self._add_records('1')
        manifest = self._create_export()
        r1 = self._download(manifest, 2, 1)
        r2 = self._download(manifest, 2, 1)
        self.assertEqual(r1.content, r2.content)
        self.assertEqual(hashlib.sha256(r1.content).hexdigest(), manifest['chunks'][2]['checksum'])

    def test_cursor_cannot_be_replayed_by_another_user(self):
        self._add_records('1')
        manifest = self._create_export()

        other_admin = User.objects.create_user('other-admin', 'testpass123', role='admin')
        other = APIClient()
        other.credentials(HTTP_AUTHORIZATION=f'Bearer {generate_token(other_admin)}')

        resp = other.get('/api/export/snapshots/manifest/', {'cursor': manifest['cursor']})
        self.assertEqual(resp.status_code, 400)
        resp = other.get(
            f'/api/export/snapshots/files/0/chunks/1/', {'cursor': manifest['cursor']}
        )
        self.assertEqual(resp.status_code, 400)

    def test_chunks_split_by_chunk_size(self):
        for i in range(5):
            self._add_records(str(i))
        manifest = self._create_export(type='stock_in', chunk_size=2)
        stock_in_file = manifest['files'][0]
        self.assertEqual(stock_in_file['rows'], 5)
        self.assertEqual(stock_in_file['chunks'], 3)
        self.assertEqual(manifest['total_rows'], 5)
        # 下载三个分块，行数分别为 2 / 2 / 1
        sizes = []
        for idx in (1, 2, 3):
            resp = self._download(manifest, 0, idx)
            self.assertEqual(resp.status_code, 200)
            wb = load_workbook(io.BytesIO(resp.content))
            sizes.append(wb.active.max_row - 1)
        self.assertEqual(sizes, [2, 2, 1])

    def test_empty_dataset_still_has_header_only_chunk(self):
        manifest = self._create_export(type='approval')
        self.assertEqual(manifest['files'][0]['rows'], 0)
        resp = self._download(manifest, 0, 1)
        self.assertEqual(resp.status_code, 200)
        wb = load_workbook(io.BytesIO(resp.content))
        self.assertEqual(wb.active.max_row, 1)  # 仅表头

    def test_expired_snapshot_is_explicitly_rejected(self):
        manifest = self._create_export()
        cursor = manifest['cursor']
        job = ExportJob.objects.get(pk=manifest['export_id'])
        job.expires_at = timezone.now() - timedelta(seconds=1)
        job.save(update_fields=['expires_at'])

        manifest_resp = self.client.get('/api/export/snapshots/manifest/', {'cursor': cursor})
        self.assertEqual(manifest_resp.status_code, 410)
        self.assertFalse(manifest_resp.json()['success'])
        self.assertIn('过期', manifest_resp.json()['message'])

        chunk_resp = self._download(manifest, 0, 1)
        self.assertEqual(chunk_resp.status_code, 410)

        # 过期后任务状态被显式标记
        job.refresh_from_db()
        self.assertEqual(job.status, ExportJob.STATUS_EXPIRED)

    def test_tampered_or_missing_cursor_rejected(self):
        manifest = self._create_export()
        good_cursor = manifest['cursor']

        # 签名被篡改
        body, _, sig = good_cursor.partition('.')
        tampered = body + '.' + ('0' if sig[0] != '0' else '1') + sig[1:]
        resp = self.client.get('/api/export/snapshots/manifest/', {'cursor': tampered})
        self.assertEqual(resp.status_code, 400)

        # 缺少游标
        resp = self.client.get('/api/export/snapshots/manifest/')
        self.assertEqual(resp.status_code, 400)

        # 游标对应的任务已不存在
        ExportJob.objects.filter(pk=manifest['export_id']).delete()
        resp = self.client.get('/api/export/snapshots/manifest/', {'cursor': good_cursor})
        self.assertEqual(resp.status_code, 404)

    def test_unknown_chunk_returns_404(self):
        manifest = self._create_export()
        resp = self._download(manifest, 0, 99)
        self.assertEqual(resp.status_code, 404)
        resp = self.client.get(
            '/api/export/snapshots/files/9/chunks/1/', {'cursor': manifest['cursor']}
        )
        self.assertEqual(resp.status_code, 404)

    def test_permission_change_blocks_new_export_but_keeps_existing_manifest(self):
        self._add_records('1')
        manifest = self._create_export()
        cursor = manifest['cursor']

        # 管理员被降级为普通用户：不能再新建导出
        self.admin.role = 'user'
        self.admin.save(update_fields=['role'])
        create_resp = self.client.post('/api/export/snapshots/', self._range_params(), format='json')
        self.assertEqual(create_resp.status_code, 403)

        # 已生成清单不被改写：仍可查询并继续下载同一快照
        resp = self.client.get('/api/export/snapshots/manifest/', {'cursor': cursor})
        self.assertEqual(resp.status_code, 200)
        same = resp.json()['data']
        self.assertEqual(same['total_rows'], manifest['total_rows'])
        self.assertEqual(same['chunks'], manifest['chunks'])

        chunk = self._download(manifest, 0, 1)
        self.assertEqual(chunk.status_code, 200)
        self.assertEqual(hashlib.sha256(chunk.content).hexdigest(), manifest['chunks'][0]['checksum'])

    def test_normal_user_cannot_create_export(self):
        normal = User.objects.create_user('plain-user', 'testpass123', role='user')
        client = APIClient()
        client.credentials(HTTP_AUTHORIZATION=f'Bearer {generate_token(normal)}')
        resp = client.post('/api/export/snapshots/', self._range_params(), format='json')
        self.assertEqual(resp.status_code, 403)

    def test_anonymous_cannot_create_or_download(self):
        anon = APIClient()
        resp = anon.post('/api/export/snapshots/', self._range_params(), format='json')
        self.assertEqual(resp.status_code, 401)
        resp = anon.get('/api/export/snapshots/manifest/', {'cursor': 'x'})
        self.assertEqual(resp.status_code, 401)

    def test_invalid_params(self):
        resp = self.client.post(
            '/api/export/snapshots/',
            {**self._range_params(), 'type': 'bogus'},
            format='json',
        )
        self.assertEqual(resp.status_code, 400)

        bad_dates = {'start_date': '2026-10-05', 'end_date': '2026-10-01'}
        resp = self.client.post('/api/export/snapshots/', bad_dates, format='json')
        self.assertEqual(resp.status_code, 400)

    def test_frozen_snapshot_excludes_records_inserted_after_creation(self):
        """直接校验快照语义：创建时物化的数据不包含之后写入的记录。"""
        self._add_records('a')
        manifest = self._create_export(type='stock_in')
        StockIn.objects.create(
            goods=self.goods, operator=self.admin, quantity=Decimal('9'), batch_no='AFTER',
        )
        resp = self._download(manifest, 0, 1)
        wb = load_workbook(io.BytesIO(resp.content))
        batch_cells = [row[4].value for row in wb.active.iter_rows(min_row=2)]
        self.assertIn('B-a', batch_cells)
        self.assertNotIn('AFTER', batch_cells)
