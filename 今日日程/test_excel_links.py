import copy
import unittest

from excel_links import resolve


def fixture():
    return {'日程事项': {'A6': 46320, 'C6': 9.5 / 24, 'D6': 11.5 / 24, 'E6': '示例事项'},
            '秋6周': {'B5': 46314, 'G5': 46319, 'H5': 46320, 'A9': '09:35—10:00', 'G9': '09:30—11:30 示例事项'},
            '秋7周': {'B5': 46321, 'A9': '09:35—10:00'},
            '月历': {'B16': 46321, 'G14': 46319, 'H14': 46320, 'H15': '09:30—11:30 示例事项'}}


class ExcelLinkTests(unittest.TestCase):
    def test_week_move_is_read_and_mirrored(self):
        result = resolve(fixture())
        self.assertEqual(result['events']['event:6']['date'], '2026-10-24')
        patches = {(item['sheet'], item['cell']): item['value'] for item in result['patches']}
        self.assertEqual(patches['日程事项', 'A6'], 46319)
        self.assertIn('示例事项', patches['月历', 'G15'])
        self.assertIsNone(patches['月历', 'H15'])

    def test_record_edit_after_week_edit_is_not_ignored(self):
        data = fixture()
        baseline = resolve(data)['state']
        data['日程事项']['A6'] = 46321
        result = resolve(data, baseline)
        self.assertEqual(result['events']['event:6']['date'], '2026-10-26')
        self.assertTrue(any(item['sheet'] == '秋7周' and item['cell'] == 'B9' for item in result['patches']))

    def test_week_time_and_title_edits(self):
        data = fixture()
        baseline = resolve(data)['state']
        data['秋6周']['G9'] = '14:30—16:00 示例改期'
        event = resolve(data, baseline)['events']['event:6']
        self.assertEqual((event['start'], event['end'], event['title']), ('14:30', '16:00', '示例改期'))

    def test_calendar_edit_is_also_read(self):
        data = fixture()
        baseline = resolve(data)['state']
        data['月历']['H15'] = '15:00—17:00 示例事项'
        self.assertEqual(resolve(data, baseline)['events']['event:6']['start'], '15:00')

    def test_repeated_read_does_not_revert_effective_date(self):
        data = fixture()
        baseline = resolve(data)['state']
        self.assertEqual(resolve(data, baseline)['events']['event:6']['date'], '2026-10-24')

    def test_one_of_duplicate_calendar_views_can_be_edited(self):
        data = fixture()
        data['月历'].update({'H24': 46320, 'H25': data['月历']['H15']})
        baseline = resolve(data)['state']
        data['月历']['H25'] = '15:00—17:00 示例事项'
        self.assertEqual(resolve(data, baseline)['events']['event:6']['start'], '15:00')

    def test_two_events_in_one_calendar_cell_keep_their_identity(self):
        data = fixture()
        data['日程事项'].update({'A7': 46320, 'C7': 18 / 24, 'D7': 20 / 24, 'E7': '另一场考试'})
        data['秋6周']['H17'] = '18:00—20:00 另一场考试'
        data['月历']['H15'] += '\n18:00—20:00 另一场考试'
        baseline = resolve(data)['state']
        events = resolve(data, baseline)['events']
        self.assertEqual(events['event:6']['title'], '示例事项')
        self.assertEqual(events['event:7']['title'], '另一场考试')

    def test_ambiguous_changes_are_not_silently_overwritten(self):
        data = fixture()
        baseline = resolve(data)['state']
        data['日程事项']['C6'] = 10 / 24
        data['秋6周']['G9'] = '14:30—16:00 示例事项'
        with self.assertRaises(ValueError):
            resolve(data, baseline)

    def test_clearing_week_event_propagates_deletion(self):
        data = fixture()
        baseline = resolve(data)['state']
        data['秋6周']['G9'] = ''
        result = resolve(data, baseline)
        self.assertNotIn('event:6', result['events'])
        self.assertTrue(any(item['sheet'] == '日程事项' and item['cell'] == 'E6' and item['value'] is None for item in result['patches']))

    def test_input_data_is_not_mutated(self):
        data = fixture()
        previous = copy.deepcopy(data)
        resolve(data)
        self.assertEqual(data, previous)

    def test_course_assignment_edits_from_both_sides(self):
        data = fixture()
        data['课内安排'] = {'A6': 46314, 'C6': '测试课程', 'D6': 8 / 24, 'E6': 9.5 / 24, 'F6': '旧作业'}
        data['秋6周'].update({'A8': '08:00—09:30', 'B8': '测试课程\n教室\n课内：旧作业'})
        baseline = resolve(data)['state']
        data['秋6周']['B8'] = '测试课程\n教室\n课内：新作业'
        patches = resolve(data, baseline)['patches']
        self.assertTrue(any(item['sheet'] == '课内安排' and item['cell'] == 'F6' and item['value'] == '新作业' for item in patches))
        data['秋6周']['B8'] = '测试课程\n教室\n课内：旧作业'
        data['课内安排']['F6'] = '小程序作业'
        patches = resolve(data, baseline)['patches']
        self.assertTrue(any(item['sheet'] == '秋6周' and item['cell'] == 'B8' and '小程序作业' in item['value'] for item in patches))

    def test_moving_a_course_updates_its_date(self):
        data = fixture()
        data['课内安排'] = {'A6': 46314, 'C6': '测试课程', 'D6': 8 / 24, 'E6': 9.5 / 24, 'F6': ''}
        data['秋6周'].update({'A8': '08:00—09:30', 'B8': '测试课程', 'C5': 46315})
        baseline = resolve(data)['state']
        data['秋6周']['B8'] = ''
        data['秋6周']['C8'] = '测试课程'
        patches = resolve(data, baseline)['patches']
        self.assertTrue(any(item['sheet'] == '课内安排' and item['cell'] == 'A6' and item['value'] == 46315 for item in patches))


if __name__ == '__main__':
    unittest.main()
