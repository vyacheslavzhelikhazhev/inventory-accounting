import hashlib
import json
import os
from pathlib import Path
import tempfile
import unittest

from runtime_bootstrap import prepare


@unittest.skipUnless(os.name == 'nt', 'Windows file-zone metadata')
class RuntimeTests(unittest.TestCase):
    def setUp(self):
        import project_tests
        project_tests.TEST_ROOT.mkdir(parents=True, exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=project_tests.TEST_ROOT)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.name = 'pythonnet/runtime/Python.Runtime.dll'
        self.dll = self.root / self.name
        self.dll.parent.mkdir(parents=True)
        self.dll.write_bytes(b'synthetic DLL fixture')
        self.marker = Path(str(self.dll) + ':Zone.Identifier')
        self.marker.write_text('[ZoneTransfer]\nZoneId=3\n')
        self.manifest({self.name: hashlib.sha256(self.dll.read_bytes()).hexdigest()})

    def manifest(self, files):
        (self.root / 'runtime_manifest.json').write_text(json.dumps({'format': 1, 'files': files}))

    def test_verified_library_only_and_repeat_start(self):
        other = self.root / 'user-document.docx'
        other.write_bytes(b'private user document')
        marker = Path(str(other) + ':Zone.Identifier')
        marker.write_text('keep this marker')
        self.assertEqual(prepare(self.root), [self.name])
        self.assertFalse(self.marker.exists())
        self.assertEqual(prepare(self.root), [])
        self.assertEqual(marker.read_text(), 'keep this marker')
        self.assertEqual(other.read_bytes(), b'private user document')

    def test_modified_library_not_unblocked(self):
        self.dll.write_bytes(b'changed')
        self.marker.write_text('[ZoneTransfer]\nZoneId=3\n')
        with self.assertRaises(ValueError):
            prepare(self.root)
        self.assertTrue(self.marker.exists())

    def test_missing_library_aborts_before_any_unblocking(self):
        self.manifest({self.name: hashlib.sha256(self.dll.read_bytes()).hexdigest(),
                       'webview/lib/missing.dll': '0' * 64})
        with self.assertRaises(ValueError):
            prepare(self.root)
        self.assertTrue(self.marker.exists())

    def test_paths_outside_package_rejected(self):
        self.manifest({'webview/lib/../../../user.dll': '0' * 64})
        with self.assertRaises(ValueError):
            prepare(self.root)
        self.assertTrue(self.marker.exists())
