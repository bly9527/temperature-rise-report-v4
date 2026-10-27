import unittest
from types import SimpleNamespace
from unittest.mock import patch

from thermal_processor import _assign_conditions, _release_empty_condition_reservations
from thermal_gui import WorkerThread


class Summary:
    def __init__(self, values=None):
        self.values = values or {}

    def Cells(self, row, column):
        return SimpleNamespace(Value=self.values.get((row, column)))


def workbook(*names):
    return SimpleNamespace(Worksheets=[SimpleNamespace(Name=n) for n in names])


def condition(index, **extra):
    return {'id': str(index), 'name': f'condition {index}', **extra}


class ConditionCapacityTests(unittest.TestCase):
    def test_only_selected_new_conditions_receive_columns(self):
        registry = [condition(i, enabled=i == 13) for i in range(1, 21)]
        state = {'conditions': {}}
        _assign_conditions(workbook('总表'), Summary(), state, registry, {'13'})
        self.assertEqual(set(state['conditions']), {'13'})
        self.assertEqual(state['conditions']['13']['column'], 4)

    def test_full_report_can_update_existing_with_unselected_extras(self):
        records = {str(i): {'name': f'condition {i}', 'column': i + 3,
                            'sheet': f'condition {i}'} for i in range(1, 13)}
        state = {'conditions': records.copy()}
        summary = Summary({(4, i + 3): f'condition {i}' for i in range(1, 13)})
        registry = [condition(i) for i in range(1, 21)]
        _assign_conditions(workbook('总表'), summary, state, registry, {'1'})
        self.assertEqual(state['conditions'], records)
        with self.assertRaisesRegex(ValueError, '12 个工况'):
            _assign_conditions(workbook('总表'), summary, state, registry, {'13'})

    def test_selected_thirteen_conditions_still_rejected(self):
        registry = [condition(i) for i in range(1, 14)]
        with self.assertRaisesRegex(ValueError, '12 个工况'):
            _assign_conditions(workbook('总表'), Summary(), {'conditions': {}},
                               registry, {c['id'] for c in registry})

    def test_old_empty_reservations_released_but_evidence_preserved(self):
        state = {'conditions': {
            str(i): {'column': i + 3, 'sheet': f'condition {i}', 'name': f'condition {i}'}
            for i in range(1, 5)}}
        summary = Summary({(4, 4): 'existing header', (5, 5): 24.5})
        _release_empty_condition_reservations(workbook('总表', 'condition 3'), summary, state)
        self.assertEqual(set(state['conditions']), {'1', '2', '3'})

    def test_legacy_same_name_retains_unselected_completed_condition(self):
        registry = [condition(2, name='same', enabled=True),
                    condition(1, name='same', enabled=False, status='已插入')]
        state = {'conditions': {}}
        _assign_conditions(workbook('总表', 'same'), Summary({(4, 4): 'same'}),
                           state, registry, {'2'})
        self.assertEqual(state['conditions']['1']['column'], 4)
        self.assertEqual(state['conditions']['2']['column'], 5)
        self.assertEqual(state['conditions']['2']['sheet'], 'same (1)')

    def test_unchecked_new_duplicate_name_does_not_allocate(self):
        registry = [condition(1, name='same', enabled=False),
                    condition(2, name='same', enabled=True)]
        state = {'conditions': {}}
        _assign_conditions(workbook('总表'), Summary(), state, registry, {'2'})
        self.assertEqual(set(state['conditions']), {'2'})

    def test_worker_error_has_no_traceback(self):
        worker = WorkerThread('unused.xlsx', [], [], [])
        errors = []
        worker.error_signal.connect(errors.append)
        with patch('thermal_gui.process_bulk_thermal_data', side_effect=ValueError('总表已满')):
            worker.run()
        self.assertEqual(errors, ['总表已满'])


if __name__ == '__main__':
    unittest.main()
