import os
import openpyxl as xl
import openpyxl.comments

# Global variables
SUMMARY_SHEET_NAME = "Summary"
TARGET_FILE_FIRST_REPORT_COL = 5


def load_id_number_list(roster_file_name, worksheet_name="MarkEntry"):
    """
    Creates a list with the id numbers of the students.

    It opens the roster file, retrieves the id numbers of all students
    and returns them as a list.

    :param roster_file_name: target file
    :param worksheet_name: worksheet where the data are.
    If not specified the default value is "MarkEntry"
    :return: a list with all id numbers as strings
    """
    wb_january = xl.load_workbook(roster_file_name)
    id_number_list = []
    ws = wb_january[worksheet_name]
    i = 1
    offset = 5
    student_id = ws.cell(row=offset, column=1).value
    while student_id is not None:
        id_number_list += [str(student_id)]
        student_id = ws.cell(row=offset+i, column=1).value
        i += 1
    return id_number_list


def find_in_list(string_list, key_word):
    """
    Gives an index number for the given element in the given list.

    :param string_list: A list
    :param key_word: element to be found
    :return: the index number of the element
    :raises ValueError: if the element is not found
    """
    return string_list.index(key_word)


def fetch_marks_for_student(filename):
    """
    Creates a dictionary with all the marks and other info from a
    student's project marksheet.

    the dictionary contains the following keys:
    - student_number
    - student_name
    - supervisor: supervisor name
    - moderator: moderator name
    - first_report
    - progress
    - executive_presentation
    - oral
    - final_report_moderator
    - final_report_supervisor
    - final_report
    - total_mark

    :param filename: project marksheet filename
    :return: a dictionary
    """
    wb = xl.load_workbook(filename, data_only=True)
    ws = wb[SUMMARY_SHEET_NAME]
    student_number = str(ws['C7'].value)
    student_name = ws["C6"].value
    student_supervisor = ws["C10"].value.strip()
    student_moderator = ws["C11"].value.strip()
    first_report_mark = ws["H6"].value
    progress_mark = ws["H7"].value
    executive_presentation_mark = ws["H8"].value
    oral_mark = ws["H9"].value
    final_report_moderator_mark = ws["E10"].value
    final_report_supervisor_mark = ws["E11"].value
    final_report = ws["H11"].value
    total_mark = ws["H12"].value
    student_results = {
        "student_number": student_number,
        "student_name": student_name,
        "supervisor": student_supervisor,
        "moderator": student_moderator,
        "first_report": first_report_mark,
        "progress": progress_mark,
        "executive_presentation": executive_presentation_mark,
        "oral": oral_mark,
        "final_report_moderator": final_report_moderator_mark,
        "final_report_supervisor": final_report_supervisor_mark,
        "final_report": final_report,
        "total_mark": total_mark
    }
    return student_results


def record_data_for_student(results, roster_file_name, list_of_students_ids):
    """

    :param results:
    :param roster_file_name:
    :param list_of_students_ids:
    :return:
    """
    wb = xl.open(roster_file_name)
    ws = wb["MarkEntry"]
    loc = find_in_list(list_of_students_ids, results["student_number"])
    offset = 5
    ws.cell(row=loc+offset, column=TARGET_FILE_FIRST_REPORT_COL).value = results["first_report"]
    ws.cell(row=loc + offset, column=6).value = results["progress"]
    ws.cell(row=loc + offset, column=7).value = results["executive_presentation"]
    ws.cell(row=loc + offset, column=8).value = results["oral"]
    ws.cell(row=loc + offset, column=9).value = results["final_report"]
    supervisor_mark = round(4 * results["final_report_supervisor"])
    moderator_mark = round(4 * results["final_report_moderator"])
    supervisor = results["supervisor"]
    moderator = results["moderator"]
    supervisor_and_moderator_marks = xl.comments.Comment(f"supervisor ({supervisor}): {supervisor_mark}, "
                                                         f"moderator ({moderator}): {moderator_mark}", "python")
    ws.cell(row=loc + offset, column=9).comment = supervisor_and_moderator_marks
    wb.save(roster_file_name)


def copy_markssheets_to_project_roster(marksheet_folder_name, roster_file_name):
    """

    :param marksheet_folder_name: lalala
    :param roster_file_name:  lalalala
    :return:
    """
    marksheets_list = os.listdir(marksheet_folder_name)
    student_numbers_in_roster = load_id_number_list(roster_file_name)
    for file in marksheets_list:
        # marksheet_folder_name = data/Mark Sheets Jan22
        # file = xyz.xlsx
        # what_we_want = data/Mark Sheets Jan22/xyz.xlsx
        file_path = os.path.join(marksheet_folder_name, file)
        marks = fetch_marks_for_student(file_path)
        try:
            record_data_for_student(marks, roster_file_name, student_numbers_in_roster)
        except:
            print(f"could not record the marks of {file}")
