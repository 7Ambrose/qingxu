import datetime as dt
import pathlib
from unittest.mock import Mock, patch

from read_plan import agenda, BASE

day = dt.date(2026, 10, 19)
serial = (day - BASE).days
workbook = Mock(spec=pathlib.Path)
workbook.stat.return_value.st_mtime_ns = 1
weekly = {
    'B5': serial, 'A67': 'Words', 'A68': 'Run',
    'B50': 'Weekly task', 'F50': serial, 'J50': serial,
    'B75': 3, 'B76': 2, 'B77': 'Monday words',
}
records = {'A15': serial, 'B15': 1, 'D15': 'Ordinary task'}
with patch('read_plan.sheets', return_value={'秋6周': weekly, '学习记录': records}):
    result = agenda(workbook, day)
    assert sum(item['id'] == 'weekly:秋6周:50' for item in result['tasks']) == 1
    regular = [item for item in result['tasks'] if item['id'].startswith(('daily:', 'habit:'))]
    assert [item['title'] for item in regular] == ['Ordinary task', 'Run', 'Monday words']
    weekly['J50'] = serial + 1
    result = agenda(workbook, day)
    assert not any(item['id'] == 'weekly:秋6周:50' for item in result['tasks'])
    following = agenda(workbook, day + dt.timedelta(days=1))
    assert next(item for item in following['habits'] if item['id'].endswith(':67'))['title'] == 'Words'
print('PASS: weekly tasks appear once on their scheduled date; daily names and order stay independent.')
