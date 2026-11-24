import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw

from powerz_title import write_titled_powerz


class PowerZTitleTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.source = Path(directory.name) / 'source.png'
        self.result = self.source.with_name('result.png')

    def test_replace_title_without_changing_plot_or_source(self):
        im = Image.new('RGB', (800, 450), '#1f1f1f')
        draw = ImageDraw.Draw(im)
        draw.text((400, 14), 'OLD PROJECT 20261123', fill='white', anchor='mm')
        draw.line((15, 28, 784, 28), fill='#503218')
        im.save(self.source)
        original = self.source.read_bytes()
        self.assertTrue(write_titled_powerz(self.source, 'C口20V放电', self.result))
        self.assertEqual(self.source.read_bytes(), original)
        with Image.open(self.result) as updated:
            self.assertEqual(im.size, updated.size)
            diff = ImageChops.difference(im, updated.convert('RGB')).getbbox()
            self.assertIsNotNone(diff)
            self.assertLessEqual(diff[3], 28)

    def test_top_readings_preserved_without_dedicated_title(self):
        im = Image.new('RGB', (600, 300), 'black')
        ImageDraw.Draw(im).text((3, 0), 'VBUS 20V   IBUS 3A', fill='yellow')
        im.save(self.source)
        self.assertFalse(write_titled_powerz(self.source, '放电', self.result))
        with Image.open(self.result) as updated:
            offset = updated.height - im.height
            self.assertGreater(offset, 0)
            content = updated.convert('RGB').crop((0, offset, im.width, updated.height))
            self.assertIsNone(ImageChops.difference(content, im).getbbox())

    def test_square_picture_preserves_all_plot_pixels(self):
        im = Image.new('RGB', (600, 300), 'black')
        ImageDraw.Draw(im).rectangle((1, 1, 599, 299), outline='lime')
        im.save(self.source)
        write_titled_powerz(self.source, 'C口9V充电', self.result, square=True)
        with Image.open(self.result) as updated:
            self.assertEqual(updated.size, (600, 600))
            top = (600 - im.height - 20) // 2 + 20
            content = updated.convert('RGB').crop((0, top, 600, top + im.height))
            self.assertIsNone(ImageChops.difference(content, im).getbbox())
