import pathlib
import tempfile
import unittest
from unittest.mock import patch

from sync_review import review_sync


class ReviewTests(unittest.TestCase):
    def test_current_values_and_missing_tasks(self):
        with tempfile.TemporaryDirectory() as fixture:
            path = pathlib.Path(fixture) / 'plan.xlsx'
            path.write_bytes(b'fixture')
            data = {'schedule': [{'id': 'class:10', 'title': 'Course', 'taskTitle': 'Excel assignment', 'done': False}],
                    'tasks': [], 'weekly': []}
            operations = [
                {'id': 'class:10', 'date': '2026-10-20', 'fields': {'title': 'Local assignment', 'done': True}},
                {'id': 'weekly:秋6周:50', 'date': '2026-10-20', 'fields': {'title': 'Deleted task'}},
                {'id': 'local:new', 'date': '2026-10-20', 'fields': {'title': 'New task'}},
            ]
            with patch('sync_review.agenda', return_value=data):
                result = review_sync(path, operations)
            self.assertIn('Excel assignment', result['entries'][0]['excel'])
            self.assertIn('Local assignment', result['entries'][0]['local'])
            self.assertFalse(result['entries'][1]['canApply'])
            self.assertTrue(result['entries'][2]['canApply'])
            self.assertEqual(result['version'], str(path.stat().st_mtime_ns))
            self.assertEqual(len(result['digest']), 64)


if __name__ == '__main__':
    unittest.main()
