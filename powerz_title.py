"""Render the detail sheet name into the Power-Z picture embedded in a report."""
import os
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFont


def _title_height(image):
    """Recognize a centered title on a flat dark header above the plot frame."""
    width, height = image.size
    if width < 160 or height < 80:
        return 0
    bg = image.getpixel((0, 0))
    if max(bg) > 85 or max(bg) - min(bg) > 12:
        return 0
    probes = [width * pct // 100 for pct in (12, 20, 80, 88)]
    for y in range(1, min(height // 8, 100)):
        if all(max(abs(a - b) for a, b in zip(image.getpixel((x, y)), bg)) > 10
               for x in probes):
            break
    else:
        return 0
    if y < 9:
        return 0
    band = image.crop((0, 0, width, y))
    difference = ImageChops.difference(band, Image.new('RGB', band.size, bg))
    ink = difference.convert('L').point(lambda n: 255 if n > 38 else 0).getbbox()
    if (ink and ink[0] >= width * .06 and ink[2] <= width * .94
            and abs(ink[0] + ink[2] - width) <= width * .08
            and 0 < ink[1] < ink[3] < y):
        return y
    return 0


def write_titled_powerz(source, title, destination, *, square=False):
    with Image.open(source) as original:
        image = original.convert('RGBA')
    width, height = image.size
    header = _title_height(image.convert('RGB'))
    replaced = bool(header)
    if replaced:
        background = image.getpixel((0, 0))
    else:
        # Cropped instrument captures can start with readings on their first row.
        header = max(20, round(width / 60))
        background = (31, 31, 31, 255)
        canvas = Image.new('RGBA', (width, height + header), background)
        canvas.paste(image, (0, header))
        image = canvas

    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, width - 1, header - 1), fill=background)
    fonts = Path(os.environ['WINDIR']) / 'Fonts'
    path = next((p for p in (fonts / 'msyh.ttc', fonts / 'simsun.ttc',
                             fonts / 'msjh.ttc') if p.is_file()), None)
    if path is None:
        raise ValueError('绘制截图标题需要微软雅黑、宋体或微软正黑体，请安装其中一种字体。')
    label = ' '.join(title.splitlines())
    size = max(1, round(header * .78))
    font = ImageFont.truetype(str(path), size)
    left, top, right, bottom = draw.textbbox((0, 0), label, font=font)
    scale = min(1, (width * .96) / max(right - left, 1),
                (header * .92) / max(bottom - top, 1))
    if scale < 1:
        font = font.font_variant(size=max(1, int(size * scale)))
        left, top, right, bottom = draw.textbbox((0, 0), label, font=font)
    draw.text(((width - right - left) // 2, (header - bottom - top) // 2),
              label, font=font, fill='white')
    if square:
        side = max(image.size)
        canvas = Image.new('RGBA', (side, side), background)
        canvas.paste(image, ((side - image.width) // 2, (side - image.height) // 2))
        image = canvas
    image.save(destination, format='PNG')
    return replaced
