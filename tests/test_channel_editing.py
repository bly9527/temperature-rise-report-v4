import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import thermal_gui as gui
from PyQt5.QtCore import Qt, QItemSelectionModel
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

    def click_save(self):
        button = next(b for b in self.window.findChildren(gui.QPushButton)
                      if '保存当前配置' in b.text())
        QTest.mouseClick(button, Qt.LeftButton)

    def test_save_button_recreates_missing_config_folder(self):
        config_dir = self.config_path.parent
        config_dir.rename(config_dir.with_name('moved_configs'))
        self.click_save()
        saved = json.loads(self.config_path.read_text(encoding='utf-8'))
        self.assertEqual(saved['conditions'][0]['id'], 'existing')

    def test_save_failure_preserves_existing_config(self):
        original = self.config_path.read_bytes()
        self.window.table_conds.item(0, 2).setText('新的工况名')
        with patch.object(gui.os, 'replace', side_effect=PermissionError('locked')):
            with patch.object(gui.QMessageBox, 'warning') as warning:
                self.click_save()
        warning.assert_called_once()
        self.assertEqual(self.config_path.read_bytes(), original)
        self.assertEqual(list(self.config_path.parent.iterdir()), [self.config_path])

    def save_preset(self, name='通用布点'):
        with patch.object(gui.QInputDialog, 'getText', return_value=(name, True)):
            self.window.save_channel_preset()

    def test_channel_preset_survives_restart_and_applies_to_new_project(self):
        channels = [
            {'hardware_name': '304 (℃)', 'enabled': True, 'custom_name': '板端NTC', 'limit': '65'},
            {'hardware_name': '320 (℃)', 'enabled': False, 'custom_name': '环境', 'limit': ''},
        ]
        self.window.apply_channel_config(channels)
        self.save_preset('通道/布点方案')
        presets = self.window.read_channel_presets()
        self.assertEqual(len(presets), 1)
        self.assertEqual(presets[0][1]['channels'], channels)
        target = self.config_path.with_name('new.json')
        target.write_text(json.dumps({'project_excel': 'new_report.xlsx', 'conditions': [
            {'id': 'new', 'name': '新工况', 'csv': 'new.csv', 'powerz': 'power.png', 'status': '待生成'}
        ], 'channels': [{'hardware_name': '101 (℃)', 'enabled': True,
                         'custom_name': '旧通道', 'limit': '30'}]}), encoding='utf-8')
        with patch.object(gui, 'get_base_dir', return_value=self.directory.name):
            reopened = gui.ThermalApp()
        try:
            reopened.cbo_projects.setCurrentText('new')
            before = reopened.collect_project_config()
            with patch.object(gui.QInputDialog, 'getItem', return_value=('通道/布点方案', True)):
                with patch.object(gui.QMessageBox, 'question', return_value=gui.QMessageBox.Yes):
                    reopened.apply_channel_preset()
            saved = json.loads(target.read_text(encoding='utf-8'))
            self.assertEqual(saved['channels'], channels)
            for key in before.keys() - {'channels'}:
                self.assertEqual(saved[key], before[key])
            reopened.on_project_changed()
            self.assertEqual(reopened.collect_channel_config(), channels)
        finally:
            reopened.close()

    def test_preset_overwrite_cancel_and_write_failure_preserve_saved_file(self):
        self.save_preset()
        path = Path(self.window.read_channel_presets()[0][0])
        original = path.read_bytes()
        self.window.table_channels.item(0, 1).setText('更改')
        with patch.object(gui.QMessageBox, 'question', return_value=gui.QMessageBox.No):
            self.save_preset()
        self.assertEqual(path.read_bytes(), original)
        with patch.object(gui.QMessageBox, 'question', return_value=gui.QMessageBox.Yes):
            with patch.object(gui.os, 'replace', side_effect=PermissionError('locked')):
                with patch.object(gui.QMessageBox, 'warning') as warning:
                    self.save_preset()
        warning.assert_called_once()
        self.assertEqual(path.read_bytes(), original)
        self.assertEqual(list(path.parent.iterdir()), [path])

    def test_invalid_preset_does_not_partially_apply_channels(self):
        self.save_preset()
        path = Path(self.window.read_channel_presets()[0][0])
        value = json.loads(path.read_text(encoding='utf-8'))
        value['channels'] = [{'hardware_name': '999 (℃)', 'enabled': True,
                              'custom_name': '错误', 'limit': ''}]
        path.write_text(json.dumps(value), encoding='utf-8')
        before = self.window.collect_project_config()
        with patch.object(gui.QMessageBox, 'warning') as warning:
            self.window.apply_channel_preset()
        warning.assert_called_once()
        self.assertEqual(self.window.collect_project_config(), before)

    def test_checkbox_still_toggles_without_editing_label(self):
        table = self.window.table_channels
        checkbox = table.cellWidget(0, 0).findChild(gui.QCheckBox)
        QTest.mouseClick(checkbox, Qt.LeftButton)
        self.assertTrue(checkbox.isChecked())
        QTest.mouseClick(checkbox, Qt.LeftButton)
        self.assertFalse(checkbox.isChecked())

    def test_multi_selection_confirmation_can_be_cancelled(self):
        self.window.add_cond_row('第二工况')
        self.window.add_cond_row('第三工况')
        table = self.window.table_conds
        selection = table.selectionModel()
        for row in (0, 2):
            selection.select(table.model().index(row, 2),
                             QItemSelectionModel.Select | QItemSelectionModel.Rows)
        with patch.object(gui.QMessageBox, 'question', return_value=gui.QMessageBox.No) as ask:
            self.window.delete_selected_conditions()
        self.assertEqual(table.rowCount(), 3)
        message = ask.call_args.args[2]
        self.assertIn('原工况', message)
        self.assertIn('第三工况', message)
        self.assertNotIn('第二工况', message)


if __name__ == '__main__':
    unittest.main()
