"""Excel presentation used by the development version of the generator."""
import math
import re
from pathlib import Path

from openpyxl.utils import get_column_letter
from PIL import Image

from thermal_phases import image_digest, locate_events, meter_steps, phase_text
from powerz_db import analyze_meter


_CHARGING_TAPER = re.compile(r'尾段渐降|恒流转恒压|持续减流')
_NOMINAL_STAGE = re.compile(r'^约?\s*(\d+(?:\.\d+)?)\s*V\s*/\s*(\d+(?:\.\d+)?)\s*A\b')


def report_stage_text(value, charging):
    lines = []
    for line in phase_text(value).splitlines():
        if charging and _CHARGING_TAPER.search(line):
            continue
        line = _NOMINAL_STAGE.sub(r'\1V/\2A', line)
        line = re.sub(r'(?<!\d)(\d+):([0-5]\d):([0-5]\d)(?![\d.:])',
                      lambda m: f'{int(m[1])}h{int(m[2])}min{int(m[3])}s', line)
        lines.append(line)
    return '\n'.join(lines)


def color_verdicts(cells, address):
    cells.Font.Color = 0
    rules = cells.FormatConditions
    for verdict, color in [('PASS', 0x008000), ('FAIL', 0x0000FF)]:
        expression = f'=RIGHT({address},4)="{verdict}"'
        if not any(str(rule.Formula1).upper() == expression for rule in rules):
            rule = rules.Add(2, 1, expression)
            rule.Font.Color = color
            rule.SetFirstPriority()


def configure_summary(summary, channels):
    summary.Name = '温升测试'
    summary.Range('A28:C28').Merge()
    summary.Range('A28').Value = '功率阶段'
    summary.Range('D28:O28').WrapText = True
    by_number = {re.search(r'\d+', ch['hardware_name']).group(): ch for ch in channels}
    positions = []
    for row in range(5, 28):
        key = re.search(r'\d+', str(summary.Cells(row, 2).Value or ''))
        if not key or key.group() not in by_number:
            continue
        channel = by_number[key.group()]
        limit = str(channel.get('limit','')).strip()
        if limit:
            try:
                maximum = float(limit)
                if not math.isfinite(maximum):
                    raise ValueError()
            except ValueError:
                raise ValueError(f'{channel["custom_name"]} 的报警限值须为数值。') from None
            summary.Cells(row, 16).Value = maximum
            summary.Cells(row, 17).Formula = (
                f'=IF(COUNT(D{row}:O{row})=0,"待测",IF(MAX(D{row}:O{row})>P{row},"FAIL","PASS"))')
        else:
            summary.Cells(row, 16).ClearContents()
            summary.Cells(row, 17).Value = '仅记录'
        positions.append((channel['hardware_name'], row))
    label = '外壳判定' if '外壳判定' in str(summary.Range('J1').Formula) else '整体判定'
    summary.Range('J1').Formula = (f'="{label}："&IF(COUNTIF(Q5:Q27,"FAIL"),"FAIL",'
                                   'IF(COUNTIF(Q5:Q27,"PASS"),"PASS","待测"))')
    color_verdicts(summary.Range('Q5:Q27'), 'Q5')
    color_verdicts(summary.Range('J1'), 'J1')
    return dict(positions)


def place_image(shape, path, left, top, width, height):
    with Image.open(path) as image:
        ratio = image.width / image.height
    shape.LockAspectRatio = 0
    shape.Width = min(width, height * ratio)
    shape.Height = shape.Width / ratio
    shape.Left, shape.Top = left, top


def format_detail(sheet, summary, column, cond, channels, positions, settings, log):
    if settings and settings.get('image_hash') != image_digest(cond['powerz']):
        raise ValueError(f'[{cond["name"]}] 截图已更换，请在“功率阶段 / 降档标记”重新确认。')
    db_path = settings.get('db_path')
    db_data = None
    if db_path:
        if not Path(db_path).is_file():
            raise ValueError(f'[{cond["name"]}] 找不到已绑定的 DB，请重新选择曲线 DB。')
        if settings.get('db_hash') and image_digest(db_path) != settings['db_hash']:
            raise ValueError(f'[{cond["name"]}] DB 内容已变化，请重新选择该文件并保存。')
        db_data = analyze_meter(db_path)
    elif not settings and cond['powerz']:
        try:
            db_data = meter_steps(cond['powerz'])
        except ValueError as error:
            log(f'[{cond["name"]}] 同名文件未作为曲线来源：{error}')
        if db_data:
            db_path = Path(cond['powerz']).with_suffix('.db')
        else:
            log(f'[{cond["name"]}] 请在“功率阶段 / 降档标记”按截图刻度选点，保存后可重复生成。')
    if db_data:
        settings = dict(settings, **db_data)
        log(f'[{cond["name"]}] 原始 DB：{Path(db_path).name}，{db_data["samples"]:,} 个采样。')
    n = len(channels)
    last = n + 2
    end = len(cond['data']) + 4
    sheet.Range(sheet.Cells(4,3),sheet.Cells(4,last)).Value = (
        tuple(ch['custom_name'] for ch in channels),)
    sheet.Range(sheet.Cells(1,3),sheet.Cells(1,last)).ClearContents()
    sheet.Cells(2,2).Value = '最高温度'
    sheet.Range(sheet.Cells(2,3),sheet.Cells(2,last)).NumberFormat = '0.00'
    sheet.Range(f'A5:A{end}').NumberFormat = 'yyyy/mm/dd hh:mm:ss'
    sheet.Columns(1).ColumnWidth = 22
    sheet.Rows(4).RowHeight = 30
    for j, channel in enumerate(channels, 3):
        row = positions[channel['hardware_name']]
        name = sheet.Name.replace("'", "''")
        summary.Cells(row,column).Formula = f"='{name}'!${get_column_letter(j)}$2"
        sheet.Cells(4,j).Interior.Color = summary.Cells(row,3).Interior.Color
        sheet.Columns(j).ColumnWidth = 13
    charging = re.search(r'放电|充电', cond['name'])
    charging = bool(charging and charging.group() == '充电')
    desc = report_stage_text(settings.get('summary', ''), charging)
    summary.Cells(28,column).Value = desc or None
    if desc:
        summary.Columns(column).ColumnWidth = max(26,summary.Columns(column).ColumnWidth)
        summary.Rows(28).RowHeight = max(summary.Rows(28).RowHeight, min(409, 26*(desc.count('\n')+1)))
    events = settings.get('events', [])
    if charging:
        events = [event for event in events if not _CHARGING_TAPER.search(event['label'])]
    if events:
        _, marks, skipped = locate_events(cond['data'], channels, events, settings.get('load_start',0))
        for offset, labels in marks.items():
            row = offset + 5
            sheet.Range(sheet.Cells(row,1), sheet.Cells(row,last+1)).Interior.Color = 65535
            sheet.Cells(row,last+1).Value = ' / '.join(labels)
        sheet.Cells(4,last+1).Value = '功率变化点（行位置估算）'
        sheet.Columns(last+1).ColumnWidth = 37
        log(f'[{cond["name"]}] {len(marks)} 行降功率标黄。' + '；'.join(skipped))
    side = max(17, last+2)
    box = sheet.Range(sheet.Cells(1,side), sheet.Cells(4,side+11))
    box.Merge()
    box.WrapText = True
    box.Value = desc or None
    box.VerticalAlignment = -4108
    if desc:
        sheet.Rows(1).RowHeight = max(24, min(320, 16*len(desc.splitlines())))
    top = sheet.Cells(6,side).Top
    left = sheet.Cells(6,side).Left
    if cond['powerz']:
        picture = sheet.Shapes.Item('ThermalPowerZ')
        picture.LockAspectRatio = 0
        picture.Width = picture.Height = 510
        picture.Left, picture.Top = left, top
        picture.Placement = 2  # Move with cells without resizing.
        picture.LockAspectRatio = -1
        top += picture.Height + 10
    r = 6
    while sheet.Cells(r,side).Top < top:
        r += 1
    photo = settings.get('detail_photo')
    if photo:
        if not Path(photo).is_file():
            raise ValueError(f'测温点位图片不存在：{photo}')
        sheet.Cells(r,side).Value = '测温点位参考'
        top = sheet.Cells(r+1,side).Top
        picture = sheet.Shapes.AddPicture(photo,0,-1,left,top,300,240)
        place_image(picture,photo,left,top,420,325)
