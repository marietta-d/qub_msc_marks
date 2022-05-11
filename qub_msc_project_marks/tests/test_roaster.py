import unittest
import roasted_roster.project_marks as roast


class MyTestCase(unittest.TestCase):

    def test_load_id_number_list(self):
        roster_filename = "../data/january_ELE_8060_2211_12915_1.xlsx"
        list_student_ids = roast.load_id_number_list(roster_filename)
        actual_length = len(list_student_ids)
        expected_length = 82
        self.assertEqual(expected_length, actual_length)
        actual_first_id = list_student_ids[0]
        expected_first_id = '40319670'
        self.assertEqual(expected_first_id, actual_first_id)
        actual_last_id = list_student_ids[-1]
        expected_last_id = '40286842'
        self.assertEqual(expected_last_id, actual_last_id)

    def test_find_in_list(self):
        my_list = ["abc", "def", "ghi"]
        keyword = "ghi"
        actual_loc = roast.find_in_list(my_list, keyword)
        expected_loc = 2
        self.assertEqual(expected_loc, actual_loc)

    def test_find_in_list_not_existing(self):
        my_list = ["abc", "def", "ghi"]
        keyword = "k"
        with self.assertRaises(Exception) as _context:
            roast.find_in_list(my_list, keyword)



    def test_fetch_marks_for_student(self):
        marksheet_filename = "../data/Mark Sheets Jan22/Dhivya.xlsx"
        marks = roast.fetch_marks_for_student(marksheet_filename)
        expected_first_report = 65
        expected_progress = 73
        expected_executive = 68
        expected_oral = 67
        expected_final = 59
        expected_student_number = "40309241"
        expected_supervisor = "Oleksandr Malyuskin"
        self.assertEqual(expected_first_report, marks["first_report"])
        self.assertEqual(expected_progress, marks["progress"])
        self.assertEqual(expected_executive, marks["executive_presentation"])
        self.assertEqual(expected_oral, marks["oral"])
        self.assertEqual(expected_final, marks["final_report"])
        self.assertEqual(expected_student_number, marks["student_number"])


if __name__ == '__main__':
    unittest.main()
