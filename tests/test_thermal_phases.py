import sqlite3
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtGui import QImage, QColor
from phase_editor import curve_steps
from thermal_phases import clock_text, locate_events, meter_steps, parse_elapsed, phase_text
from powerz_db import analyze_meter
from dev_report import report_stage_text


class PhaseTests(unittest.TestCase):
    channels = [{'custom_name': '面壳'}, {'custom_name': '环温'}]

    def test_report_power_stage_format_omits_charging_taper(self):
        source = '约9V / 1.55A 0:47:10\n约9V / 尾段渐降 0:52:50\n约5V / 1.5A 1:00:02'
        self.assertEqual(report_stage_text(source, True), '9V/1.55A 0h47min10s\n5V/1.5A 1h0min2s')
        self.assertIn('尾段渐降', report_stage_text(source, False))

    def data(self):
        return [(46000 + i*5/86400, i, 25+max(0,min(i-12,12))*.3, 25+i*.5)
                for i in range(45)]

    def test_temperature_onset_and_nearest_scan(self):
        origin, marks, skipped = locate_events(self.data(), self.channels, [
            {'seconds': 25, 'label':'降至 5V'}, {'seconds':250,'label':'停载'}], 5)
        self.assertEqual(origin, 12)
        self.assertEqual(marks, {16:['0:00:25 降至 5V']})
        self.assertEqual(len(skipped), 1)

    def test_no_interpolation_through_a_hole(self):
        rows = self.data()
        del rows[23:30]
        _,marks,skipped = locate_events(rows,self.channels,[{'seconds':60,'label':'降档'}])
        self.assertFalse(marks)
        self.assertTrue(skipped)

    def test_ambient_or_one_spike_does_not_anchor(self):
        rows=[(r[0],r[1],30 if i==8 else 25,r[3]) for i,r in enumerate(self.data())]
        self.assertIsNone(locate_events(rows,self.channels,[])[0])
        self.assertIsNone(locate_events(rows,[{'custom_name':'环境温度'}],[])[0])

    def test_summary(self):
        self.assertEqual(phase_text('测温采集 1:23:45\n仪器记录 1:22:00\n带载阶段时长待核'),
                         '仪器记录 1:22:00')
        with self.assertRaises(ValueError):
            parse_elapsed('0:66:30')

    def test_meter_finds_reduction_in_voltage_at_higher_current(self):
        with tempfile.TemporaryDirectory() as folder:
            db=Path(folder)/'测试.db'
            with sqlite3.connect(db) as c:
                c.execute('create table table_1(Unix,VBUS,IBUS)')
                c.executemany('insert into table_1 values (?,?,?)',
                              [(i,9 if i<30 else 5,1.1 if i<30 else 1.5) for i in range(60)])
            c.close()
            out=meter_steps(db.with_suffix('.png'))
            self.assertEqual(out['events'], [{'seconds':30,'label':'降至约 5.00V / 1.50A'}])

    def test_image_candidates(self):
        img=QImage(250,120,QImage.Format_RGB32)
        img.fill(QColor('black'))
        for x in range(10,240):
            img.setPixelColor(x, 35 if x<100 else 65, QColor('cyan'))
        self.assertEqual(curve_steps(img,10,239,50,35),[100])

    def test_fractional_timestamp_and_offset_from_original_recording(self):
        with tempfile.TemporaryDirectory() as folder:
            db = Path(folder) / '没有同名截图.db'
            c = sqlite3.connect(db)
            with c:
                c.execute('CREATE TABLE table_1(Unix, ElapsedTime, VBUS, IBUS)')
                c.executemany('INSERT INTO table_1 VALUES (?,?,?,?)', [
                    (1700000000+i/10,clock_text(60+i/10),9,2 if i<213 else 1)
                    for i in range(500)])
            c.close()
            result = analyze_meter(db)
            self.assertEqual(result['load_start'],60)
            self.assertEqual(result['events'][0]['seconds'],81.3)
            self.assertEqual(result['event_table'][1][-3:],[9.0,1.0,9.0])
            self.assertEqual(parse_elapsed(clock_text(81.3)),81.3)

    def test_missing_or_incompatible_database(self):
        with tempfile.TemporaryDirectory() as folder:
            db = Path(folder) / 'missing.db'
            with self.assertRaises(ValueError):
                analyze_meter(db)
            self.assertFalse(db.exists())
            db.write_bytes(b'unrelated file')
            with self.assertRaises(ValueError):
                analyze_meter(db)
