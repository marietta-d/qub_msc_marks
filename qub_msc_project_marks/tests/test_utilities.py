import unittest
import roasted_roster.utilities as util


class UtilitiesTestCase(unittest.TestCase):
    def test_find_in_list(self):
        my_list = ["abc", "def", "ghi"]
        keyword = "ghi"
        actual_loc = util.find_in_list(my_list, keyword)
        expected_loc = 2
        self.assertEqual(expected_loc, actual_loc)

    def test_find_in_list_not_there(self):
        my_list = ["abc", "def", "ghi"]
        keyword = "h"
        actual_loc = util.find_in_list(my_list, keyword)
        expected_loc = None
        self.assertEqual(expected_loc, actual_loc)


if __name__ == '__main__':
    unittest.main()
