import json
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from thermal_processor import (
    _condition_result_to_delete, _compact_condition_columns,
    _commit_condition_deletion, _validate_condition_positions,
)


class Summary:
    def __init__(self, values):
        self.values = dict(values)

    def Cells(self, row, column):
        return SimpleNamespace(Value=self.values.get((row, column)), row=row, column=column)

    def Range(self, start, end):
        summary = self
        class Range:
            def __init__(self):
                self.coordinates = [(r, c) for r in range(start.row, end.row + 1)
                                    for c in range(start.column, end.column + 1)]

            def Copy(self, target):
                values = [summary.values.get(p) for p in self.coordinates]
                for p, value in zip(target.coordinates, values):
                    summary.values[p] = value

            def ClearContents(self):
                for p in self.coordinates:
                    summary.values.pop(p, None)
        return Range()


class DeletionTests(unittest.TestCase):
    def test_compaction_preserves_unmanaged_results_and_protected_areas(self):
        summary = Summary({(4, 4): 'A', (5, 4): 20, (4, 6): 'B', (5, 6): 30,
                           (4, 9): 'manual', (5, 9): 40,
                           (5, 16): 'spec', (5, 17): 'pass', (30, 4): 'photo'})
        state = {'conditions': {'a': {'column': 4}, 'b': {'column': 6}}}
        _compact_condition_columns(summary, state)
        self.assertEqual([summary.Cells(4, c).Value for c in (4, 5, 6)], ['A', 'B', 'manual'])
        self.assertEqual([summary.Cells(5, c).Value for c in (4, 5, 6)], [20, 30, 40])
        self.assertIsNone(summary.Cells(4, 9).Value)
        self.assertEqual(state['conditions']['b']['column'], 5)
        self.assertEqual(summary.Cells(5, 16).Value, 'spec')
        self.assertEqual(summary.Cells(5, 17).Value, 'pass')
        self.assertEqual(summary.Cells(30, 4).Value, 'photo')

    def test_delete_is_by_id_even_when_ui_name_was_edited(self):
        record = {'name': 'old', 'column': 5, 'sheet': 'old (1)'}
        state = {'conditions': {'id2': record}}
        self.assertEqual(_condition_result_to_delete(
            Summary({(4, 5): 'old'}), state, {'id': 'id2', 'name': 'edited'}, []), ('id2', record))

    def test_new_same_name_does_not_delete_an_existing_result(self):
        state = {'conditions': {'old': {'name': 'same', 'column': 4, 'sheet': 'same'}}}
        self.assertEqual(_condition_result_to_delete(
            Summary({(4, 4): 'same'}), state, {'id': 'new', 'name': 'same'}, []), ('new', None))

    def test_ambiguous_legacy_results_are_not_deleted(self):
        with self.assertRaisesRegex(ValueError, '重名'):
            _condition_result_to_delete(Summary({(4, 4): 'same', (4, 5): 'same'}),
                                        {'conditions': {}}, {'id': 'x', 'name': 'same'}, [])

    def test_mismatched_header_is_not_deleted(self):
        with self.assertRaisesRegex(ValueError, '不一致'):
            _condition_result_to_delete(Summary({(4, 4): 'other'}),
                {'conditions': {'x': {'name': 'expected', 'column': 4, 'sheet': 'expected'}}},
                {'id': 'x', 'name': 'expected'}, [])

    def test_summary_sheet_cannot_be_a_deletion_target(self):
        with self.assertRaisesRegex(ValueError, '位置记录异常'):
            _validate_condition_positions({'conditions': {'x': {'column': 4, 'sheet': '总表'}}})

    def test_config_commit_failure_rolls_back_report(self):
        with tempfile.TemporaryDirectory() as d:
            report, work, config = [Path(d)/n for n in ('report.xlsx', 'work.xlsx', 'config.json')]
            report.write_bytes(b'original report'); work.write_bytes(b'changed report')
            config.write_text('{"old":true}')
            replace = os.replace
            def fail_config(source, target):
                if str(target) == str(config):
                    raise PermissionError('config locked')
                return replace(source, target)
            with patch('thermal_processor.os.replace', side_effect=fail_config):
                with self.assertRaisesRegex(OSError, '配置保存失败'):
                    _commit_condition_deletion(str(work), str(report), str(config), {'new': True})
            self.assertEqual(report.read_bytes(), b'original report')
            self.assertEqual(json.loads(config.read_text()), {'old': True})
            self.assertFalse(list(Path(d).glob('~condition_*')))

    def test_locked_report_keeps_config(self):
        with tempfile.TemporaryDirectory() as d:
            report, work, config = [Path(d)/n for n in ('report.xlsx', 'work.xlsx', 'config.json')]
            report.write_bytes(b'old'); work.write_bytes(b'new'); config.write_bytes(b'old config')
            with patch('thermal_processor.os.replace', side_effect=PermissionError('locked')):
                with self.assertRaisesRegex(PermissionError, '关闭 Excel/WPS'):
                    _commit_condition_deletion(str(work), str(report), str(config), {})
            self.assertEqual(report.read_bytes(), b'old')
            self.assertEqual(config.read_bytes(), b'old config')


if __name__ == '__main__':
    unittest.main()
