"""Development-only Power-Z image annotation dialog."""
from statistics import median
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
                             QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPlainTextEdit,
                             QPushButton, QVBoxLayout)
from thermal_phases import clock_text, image_digest, parse_elapsed, phase_text
from powerz_db import POWERZ_DB_FILTER, bind_meter, switch_to_picture


def curve_steps(image, x0, x1, x, y):
    """Suggest sustained downward steps on the clicked trace between the ticks."""
    color = image.pixelColor(x, y)
    rgb = color.getRgb()[:3]
    if max(rgb) - min(rgb) < 25:
        raise ValueError('请点在有颜色的电流或功率曲线上。')
    heights = []
    for col in range(x0, x1 + 1):
        rows = [r for r in range(int(image.height() * .04), int(image.height() * .93))
                if max(abs(a-b) for a,b in zip(image.pixelColor(col,r).getRgb()[:3], rgb)) < 32]
        heights.append(median(rows) if rows else None)
    reach = max(3, len(heights) // 170)
    threshold = max(4, image.height() * .017)
    hits = []
    for pos in range(reach, len(heights) - reach):
        before = [v for v in heights[pos-reach:pos] if v is not None]
        after = [v for v in heights[pos:pos+reach] if v is not None]
        if min(len(before), len(after)) < reach * .6:
            continue
        change = median(after) - median(before)
        if change < threshold:
            continue
        # Locate the first pixel on the lower plateau rather than the leading window.
        if heights[pos] is None or heights[pos] < median(before) + change / 2:
            continue
        if not hits or pos - hits[-1] > reach:
            hits.append(pos)
    return [x0 + idx for idx in hits]


class PhaseEditor(QDialog):
    def __init__(self, picture, settings, parent):
        super().__init__(parent)
        self.setWindowTitle('功率阶段与降档标记')
        self.digest = image_digest(picture)
        self.picture = picture
        if settings.get('db_path') or settings.get('image_hash') == self.digest:
            self.settings = dict(settings)
        else:
            self.settings = {}
            if settings.get('detail_photo'):
                self.settings['detail_photo'] = settings['detail_photo']
        layout = QVBoxLayout(self)
        db_tools = QHBoxLayout()
        self.source = QLabel('来源：截图选点')
        db_tools.addWidget(self.source)
        choose = QPushButton('选择曲线 DB')
        choose.clicked.connect(self.choose_db)
        db_tools.addWidget(choose)
        manual = QPushButton('使用截图选点')
        manual.clicked.connect(self.use_picture)
        db_tools.addWidget(manual)
        layout.addLayout(db_tools)
        pix = QPixmap(picture).scaled(870, 370, Qt.KeepAspectRatio, Qt.SmoothTransformation) if picture else QPixmap()
        self.preview = QLabel()
        if not pix.isNull():
            self.preview.setPixmap(pix)
            self.preview.setFixedSize(pix.size())
            self.preview.mousePressEvent = self.pick
        layout.addWidget(self.preview)
        form = QFormLayout()
        layout.addLayout(form)
        self.mode = QComboBox()
        self.mode.addItems(['点图设置横轴零点','点图设置第二个刻度','点图添加降档点','点曲线识别降档候选'])
        self.end_tick = QLineEdit()
        self.end_tick.setPlaceholderText('第二个刻度的图上时间，例如 2:00:00')
        self.load = QLineEdit(clock_text(self.settings.get('load_start', 0)))
        form.addRow('点图操作', self.mode)
        form.addRow('第二个刻度', self.end_tick)
        form.addRow('记录内开始带载时间', self.load)
        self.events = QPlainTextEdit('\n'.join(clock_text(e['seconds'])+' '+e['label']
                                                for e in self.settings.get('events', [])))
        self.events.setPlaceholderText('每行一个，可点图或直接填，例如 0:06:45 降至 20V / 2.25A')
        self.events.setMaximumHeight(90)
        form.addRow('降档时间及说明', self.events)
        self.note = QPlainTextEdit(phase_text(self.settings.get('summary','')))
        self.note.setMaximumHeight(72)
        form.addRow('总表功率阶段', self.note)
        self.message = QLabel('以测点持续升温的起点对应图上开始带载时间。请核对候选点，再保存。')
        layout.addWidget(self.message)
        self.ticks = [None, None]
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        stored = self.settings.get('db_path')
        sibling = Path(picture).with_suffix('.db') if picture else None
        if stored:
            self.show_db(stored)
        elif not settings and sibling and sibling.is_file():
            self.import_db(sibling)

    def show_db(self, path):
        self.source.setText('来源：' + Path(path).name)
        self.source.setToolTip(str(path))
        self.events.setReadOnly(True)
        self.load.setReadOnly(True)
        self.mode.setEnabled(False)
        self.end_tick.setEnabled(False)

    def choose_db(self):
        start = self.settings.get('db_path') or self.picture or ''
        path, _ = QFileDialog.getOpenFileName(self,'选择本工况的曲线数据库',
                                              str(Path(start).parent), POWERZ_DB_FILTER)
        if path:
            self.import_db(Path(path))

    def import_db(self, path):
        try:
            updated, result = bind_meter(path, self.settings)
        except (ValueError,OSError) as error:
            self.message.setText(str(error))
            return
        if updated['db_hash'] != self.settings.get('db_hash'):
            self.note.setPlainText(phase_text(updated.get('summary', '')))
        self.settings = updated
        self.events.setPlainText('\n'.join(clock_text(e['seconds'])+' '+e['label'] for e in result['events']))
        self.load.setText(clock_text(result['load_start']))
        self.message.setText(f'已读取 {result["samples"]:,} 个采样，典型间隔 {result["period"]:.3g} 秒，'
                             f'检出 {len(result["events"])} 个变化点。')
        self.show_db(path)

    def use_picture(self):
        if self.settings.get('db_path'):
            self.events.clear()
            self.load.setText('0:00:00')
            self.note.clear()
        self.settings = switch_to_picture(self.settings)
        self.source.setText('来源：截图选点')
        self.source.setToolTip('')
        self.events.setReadOnly(False)
        self.load.setReadOnly(False)
        self.mode.setEnabled(True)
        self.end_tick.setEnabled(True)

    def pick(self, click):
        if not self.mode.isEnabled():
            return
        mode, x = self.mode.currentIndex(), click.x()
        try:
            if mode < 2:
                if mode == 1:
                    parse_elapsed(self.end_tick.text())
                    if self.ticks[0] is None or x <= self.ticks[0]:
                        raise ValueError('第二个刻度须在零点右侧。')
                self.ticks[mode] = x
                self.mode.setCurrentIndex(mode + 1)
                return
            x0, x1 = self.ticks
            if x0 is None or x1 is None or x1 <= x0:
                raise ValueError('请先在图上点选零点和第二个刻度。')
            span = parse_elapsed(self.end_tick.text())
            if span <= 0 or x < x0:
                raise ValueError('请检查刻度时间和降档位置。')
            points = ([x] if mode == 2 else
                      curve_steps(self.preview.pixmap().toImage(),x0,x1,x,click.y()))
            for p in points:
                self.events.appendPlainText(clock_text((p-x0)*span/(x1-x0))+' 降功率')
            self.message.setText(f'已添加 {len(points)} 个点，可编辑时间、补充档位或删除误判。')
        except ValueError as exc:
            self.message.setText(str(exc))

    def save(self):
        try:
            load = parse_elapsed(self.load.text())
            events = []
            for line in self.events.toPlainText().splitlines():
                if line.strip():
                    parts = line.strip().split(maxsplit=1)
                    if len(parts) != 2:
                        raise ValueError('每行填写一个图上时间和降档说明。')
                    seconds = parse_elapsed(parts[0])
                    if seconds < load:
                        raise ValueError('降档点早于开始带载时间，请核对。')
                    events.append({'seconds':seconds,'label':parts[1]})
            summary = phase_text(self.note.toPlainText())
            if summary != self.settings.get('summary',''):
                self.settings.pop('table', None)
            self.settings.update(image_hash=self.digest,load_start=load,events=events,summary=summary)
            self.accept()
        except ValueError as exc:
            QMessageBox.warning(self, '时间格式', str(exc))
