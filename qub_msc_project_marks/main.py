import roasted_roster.project_marks as roast

# Define the roster file, in this case the target file
roster_filename = "data/january_ELE_8060_2211_12915_1.xlsx"

# With the next line the program reads the project marks of all students
# and copies them into the target file
roast.copy_markssheets_to_project_roster("data/Mark Sheets Jan22", roster_filename)


