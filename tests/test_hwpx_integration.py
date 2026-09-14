import hashlib
from pathlib import Path
import runpy
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch

api = runpy.run_path(str(Path(__file__).resolve().parents[1] / 'tools/hwpx/check_integration.py'))


class IntegrationContractTests(unittest.TestCase):
    def test_published_without_report_is_failure(self):
        with self.assertRaises(ValueError):
            api['check_build_report']({'ok': True, 'published': True, 'report_saved': False})

    def test_draft_is_not_a_complete_smoke_build(self):
        with self.assertRaises(ValueError):
            api['check_build_report']({'ok': True, 'published': True, 'report_saved': True, 'draft': True})

    def test_complete_build(self):
        api['check_build_report']({'ok': True, 'published': True, 'report_saved': True, 'status': 'PASS'})

    def test_render_evidence_must_match_document(self):
        with tempfile.TemporaryDirectory() as tmp:
            document = Path(tmp) / 'letter.hwpx'
            image = Path(tmp) / 'page.png'
            document.write_bytes(b'fixture')
            image.write_bytes(b'fixture')
            report = {'ok': True, 'page_images': [str(image)],
                      'source_sha256': hashlib.sha256(document.read_bytes()).hexdigest()}
            api['check_render_report'](report, document)
            document.write_bytes(b'changed')
            with self.assertRaises(ValueError):
                api['check_render_report'](report, document)

    def test_missing_skill_fails_before_creating_artifacts(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(ValueError):
                api['check_integration'](Path(tmp))

    def test_wrong_commit_and_changed_guide_fail_before_build(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'scripts').mkdir()
            (root / 'references').mkdir()
            (root / 'SKILL.md').write_text('references/reference-official-letter.md', encoding='utf-8')
            (root / 'references/reference-official-letter.md').write_text('changed', encoding='utf-8')
            for name in ['one_shot.py', 'render_hwpx.py']:
                (root / 'scripts' / name).write_text('', encoding='utf-8')
            with patch('subprocess.run', return_value=SimpleNamespace(stdout='wrong-commit')):
                with self.assertRaisesRegex(ValueError, 'commit differs'):
                    api['check_integration'](root)
            import json
            contract = json.loads((api['ROOT'] / 'tools/hwpx/skill-lock.json').read_text(encoding='utf-8'))
            with patch('subprocess.run', return_value=SimpleNamespace(stdout=contract['commit'])):
                with self.assertRaisesRegex(ValueError, 'guide differs'):
                    api['check_integration'](root)


if __name__ == '__main__':
    unittest.main()
