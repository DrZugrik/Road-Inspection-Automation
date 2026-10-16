import unittest
from pathlib import Path
from types import SimpleNamespace

from run_manifest import build_run_manifest


class RunManifestTests(unittest.TestCase):
    def test_summary_counts_folder_statuses(self):
        config = SimpleNamespace(template_path=Path('template.docx'), catalog_path=Path('catalog.xlsx'), output_dir=Path('output'), update_word_fields=False)
        results = [{'status': 'success'}, {'status': 'completed_with_warnings'}, {'status': 'failed'}]
        manifest = build_run_manifest(config, results)
        self.assertEqual(manifest['summary'], {'folders_total': 3, 'success': 1, 'completed_with_warnings': 1, 'failed': 1})