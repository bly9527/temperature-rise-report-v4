import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import thermal_gui as gui
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QInputMethodEvent
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QLineEdit


class ChannelEditingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        gui.configure_qt_plugins()
        cls.app = gui.QApplication.instance() or gui.QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.config_path = Path(self.directory.name) / 'configs' / 'test.json'
        self.config_path.parent.mkdir()
        self.config_path.write_text(json.dumps({
            'project_excel': 'unused.xlsx',
            'conditions': [{'id': 'existing', 'name': '原工况'}],
        }), encoding='utf-8')
        with patch.object(gui, 'get_base_dir', return_value=self.directory.name):
            self.window = gui.ThermalApp()
        self.window.show()
        QTest.qWait(50)
        self.addCleanup(self.window.close)

    def click_and_type(self, table, row, column, text):
        item = table.item(row, column)
        table.scrollToItem(item)
        point = table.visualItemRect(item).center()
        target = table.viewport().childAt(point) or table.viewport()
        QTest.mouseClick(target, Qt.LeftButton,
                         pos=target.mapFrom(table.viewport(), point))
        QTest.qWait(50)
        # The editor must exist BEFORE the first IME event. Sending directly
        # to an unedited Qt5 table with embedded widgets crashes on Windows IME.
        editor = table.focusWidget()
        self.assertIsInstance(editor, QLineEdit)
        editor.selectAll()
        self.app.sendEvent(editor, QInputMethodEvent('pinyin', []))
        event = QInputMethodEvent()
        event.setCommitString(text)
        self.app.sendEvent(editor, event)
        QTest.keyClick(editor, Qt.Key_Tab)
        self.assertEqual(item.text(), text)

    def test_single_click_chinese_name_limit_and_save_reload(self):
        table = self.window.table_channels
        for row in (0, 19, 20, 39, 40, 50, 59):
            self.click_and_type(table, row, 1, f'电芯NTC{row}')
            self.click_and_type(table, row, 2, '65.5')
        saved = self.window.save_project_config()
        self.window.on_project_changed()
        self.assertEqual(self.window.collect_project_config(), saved)
        self.assertEqual(json.loads(self.config_path.read_text(encoding='utf-8')), saved)

    def test_condition_name_uses_same_safe_editing_path(self):
        self.click_and_type(self.window.table_conds, 0, 2, '无线充电')
        saved = self.window.save_project_config()
        self.assertEqual(saved['conditions'][0]['name'], '无线充电')
        self.assertEqual(saved['conditions'][0]['id'], 'existing')

    def test_checkbox_still_toggles_without_editing_label(self):
        table = self.window.table_channels
        checkbox = table.cellWidget(0, 0).findChild(gui.QCheckBox)
        QTest.mouseClick(checkbox, Qt.LeftButton)
        self.assertTrue(checkbox.isChecked())
        QTest.mouseClick(checkbox, Qt.LeftButton)
        self.assertFalse(checkbox.isChecked())


if __name__ == '__main__':
    unittest.main()
