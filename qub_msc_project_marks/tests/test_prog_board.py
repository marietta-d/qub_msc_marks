import unittest
import roasted_roster.prog_board as pb


class ProgBoardTestCase(unittest.TestCase):
    def test_load_subject_list(self):
        board_file_name = "../data/subject_data/programme_board_template (target file blank).xlsx"
        y = pb.load_subject_list(board_file_name)
        print(y)

    def test_copy_from_marksheet_to_prog_board(self):
        marksheet_file_name = "../data/subject_data/subjects/ele8078.xlsx"
        pb_filename = "../data/subject_data/programme_board_template (target file blank).xlsx"
        subjects = pb.load_subject_list(pb_filename)
        bookkeeping_list = pb.copy_from_marksheet_to_prog_board(marksheet_file_name, pb_filename, [], subjects)
        print(bookkeeping_list)

    def test_fill_prog_board(self):
        prog_board_file_name = "../data/subject_data/programme_board_template (target file blank).xlsx"
        subject_folder = "../data/subject_data/subjects"
        final_file = pb.fill_prog_board(prog_board_file_name, subject_folder)
        print(final_file)




if __name__ == '__main__':
    unittest.main()
