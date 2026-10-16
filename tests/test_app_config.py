import argparse
import tempfile
import unittest
from pathlib import Path

from app_config import ConfigurationError, load_runtime_config


def make_args(**overrides):
    values = {
        'config': None, 'template': None, 'catalog': None, 'source': None,
        'folders': None, 'output_dir': None, 'process_subfolders': False,
        'skip_com_update': False, 'log_level': None,
    }
    values.update(overrides)
    return argparse.Namespace(**values)


class RuntimeConfigTests(unittest.TestCase):
    def test_cli_values_are_loaded(self):
        args = make_args(template=Path('template.docx'), catalog=Path('catalog.xlsx'), source=Path('object'), output_dir=Path('output'))
        config = load_runtime_config(args)
        self.assertEqual(config.template_path, Path('template.docx'))
        self.assertEqual(config.folders_to_process, [Path('object')])

    def test_subfolders_are_selected_and_sorted(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'Б').mkdir()
            (root / 'А').mkdir()
            config = load_runtime_config(make_args(template=Path('t.docx'), catalog=Path('c.xlsx'), source=root, output_dir=Path('out'), process_subfolders=True))
            self.assertEqual([folder.name for folder in config.folders_to_process], ['А', 'Б'])

    def test_required_input_is_validated(self):
        with self.assertRaises(ConfigurationError):
            load_runtime_config(make_args())