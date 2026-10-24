"""Synthetic pixel fixtures test cropping; they are not report run evidence."""
import importlib.util
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from PIL import Image

SCRIPT = Path(__file__).with_name('trim_console_shot.py')
spec = importlib.util.spec_from_file_location('trim_console_shot', SCRIPT)
trim = importlib.util.module_from_spec(spec)
spec.loader.exec_module(trim)


class TrimTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / 'raw.png'
        image = Image.new('RGB', (100, 240), (24, 24, 24))
        # The final pixel lies on an odd column, missed by the original step=2 scan.
        for x in range(1, 17, 2):
            image.putpixel((x, 170), (255, 255, 255))
        image.save(self.source)
        self.original = self.source.read_bytes()

    def tearDown(self):
        self.temp.cleanup()

    def run_cli(self, *args):
        return subprocess.run([sys.executable, str(SCRIPT), str(self.source), *map(str, args)], capture_output=True, text=True)

    def test_full_scan_keeps_final_odd_column_pixel(self):
        with Image.open(self.source) as image:
            self.assertEqual(trim.last_text_row(image, (24, 24, 24), 60, 1), 170)

    def test_default_destination_preserves_raw_and_exact_kept_pixels(self):
        self.assertEqual(self.run_cli().returncode, 0)
        self.assertEqual(self.source.read_bytes(), self.original)
        with Image.open(self.source.with_name('raw-trimmed.png')) as image, Image.open(self.source) as raw:
            self.assertEqual(image.size, (100, 187))
            self.assertEqual(image.tobytes(), raw.crop((0, 0, 100, 187)).tobytes())

    def test_refuses_in_place_write(self):
        self.assertNotEqual(self.run_cli(self.source).returncode, 0)
        self.assertEqual(self.source.read_bytes(), self.original)

    def test_refuses_hard_link_to_original(self):
        link = self.root / 'linked.png'
        try:
            os.link(self.source, link)
        except OSError as error:
            self.skipTest(str(error))
        self.assertNotEqual(self.run_cli(link).returncode, 0)
        self.assertEqual(self.source.read_bytes(), self.original)

    def test_blank_image_fails_without_output(self):
        Image.new('RGB', (100, 240), (24, 24, 24)).save(self.source)
        self.assertNotEqual(self.run_cli().returncode, 0)
        self.assertFalse(self.source.with_name('raw-trimmed.png').exists())

    def test_invalid_parameters_do_not_write_output(self):
        for args in [('--margin', '-1'), ('--inset', '50'), ('--inset', '-1'), ('--min-pixels', '0'), ('--min-pixels', '101'), ('--min-height', '0'), ('--threshold', '766')]:
            with self.subTest(args=args):
                self.assertNotEqual(self.run_cli(*args).returncode, 0)
                self.assertFalse(self.source.with_name('raw-trimmed.png').exists())


if __name__ == '__main__':
    unittest.main()
