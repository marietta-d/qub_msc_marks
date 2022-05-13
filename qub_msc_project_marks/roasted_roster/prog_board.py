import os
import openpyxl as xl
import openpyxl.comments
import roasted_roster.utilities as util


def load_subject_list(prog_board_file_name, worksheet_name="Main"):
    """
    Creates a list with the subjects.

    :param prog_board_file_name: target file
    :param worksheet_name: worksheet where the data are.
    If not specified the default value is "Main"
    :return: a list with all subjects as string
    """
    wb_target = xl.load_workbook(prog_board_file_name)
    subject_list = []
    ws = wb_target[worksheet_name]
    offset = 4
    for i in range(8):
        subject = ws.cell(row=5, column=offset+2*i).value
        subject_list += [subject]
    return subject_list


def copy_from_marksheet_to_prog_board(marksheet_file_name,
                                      prog_board_file_name,
                                      student_id_list,
                                      subjects):
    prog_board_wb = xl.load_workbook(prog_board_file_name)
    prog_board_ws = prog_board_wb["Main"]
    marksheet_wb = xl.load_workbook(marksheet_file_name, data_only=True)
    marksheet_ws = marksheet_wb["SubjectBoard"]
    module_code = marksheet_ws["C1"].value
    module_index = util.find_in_list(subjects, module_code)
    if module_index is None:
        raise Exception("wtf")
    i = 0
    row_offset_pb = 8
    column_offset_pb = 4
    while True:
        id = marksheet_ws.cell(row=5+i, column=1).value
        name = marksheet_ws.cell(row=5+i, column=2).value
        mark = marksheet_ws.cell(row=5+i, column=3).value
        result = marksheet_ws.cell(row=5+i, column=4).value
        i += 1
        if id is None:
            break
        position_in_list = util.find_in_list(student_id_list, id)
        is_new_student = False
        if position_in_list is None:
            is_new_student = True
            student_id_list += [id]
            position_in_list = len(student_id_list) - 1

        row = row_offset_pb + position_in_list
        column = column_offset_pb + 2 * module_index

        if is_new_student:
            prog_board_ws.cell(row=row, column=2).value = id
            prog_board_ws.cell(row=row, column=3).value = name

        prog_board_ws.cell(row=row, column=column).value = mark
    prog_board_wb.save(prog_board_file_name)






    return student_id_list






def fetch_grades_for_student(filename):
    """
    Creates a dictionary with all the students and other info from a
    subjectboard marksheet.

    the dictionary contains the following keys:
    - student_name
    - grade_input
    - exam_board_note

       :param filename : subject_board_file
       :return: a dictionary
    """
    wb = xl.load_workbook(filename, data_only=True)
    ws = wb["SubjectBoard"]
    student_name = ws[cell(row=4, column=2)].value.strip()
    grade_input = ws[cell(row=4, column=3)].value
    exam_board_note = ws[cell(row=4, column=4)].value.strip()
    student_results = {
        "student_name": student_name,
        "grade_input": grade_input,
        "exam_board_note": exam_board_note,
    }
    return student_results


def fetch_marks_for_subject(filename):
    """
    Creates a dictionary with all the students and other info from a
    subjectboard marksheet.

    the dictionary contains the following keys:
    - subject
    - student_id : is a dictionary

    :param filename: project marksheet filename
    :return: a dictionary
    """
    wb = xl.load_workbook(filename, data_only=True)
    ws = wb["SubjectBoard"]
    subject = ws[cell(row=1, column=3)].value
    student_id = fetch_marks_for_subject(filename)
    subject_results = {
        "subject": subject,
        "student_id": student_id,
    }
    return subject_results


