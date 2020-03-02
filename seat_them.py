from openpyxl import load_workbook
from collections import namedtuple

student_filename = 'student.xlsx'
seat_filename = 'seat.xlsx'
new_seat_filename = 'new_seat.xlsx'

Student = namedtuple('Student', ('name', 'id'))
stu_wb = load_workbook(student_filename)
stu_ws = stu_wb.active
students = []
for i in range(2, 500):
    name = stu_ws.cell(row=i, column=1).value
    idnum = stu_ws.cell(row=i, column=2).value
    if name == None:
        break
    students.append(Student(name, idnum))

wb = load_workbook(seat_filename)
ws = wb.active
idx = 0
for i in range(1, 100):
    for j in range(1, 100):
        c = ws.cell(row=i, column=j)
        if c.value == 'X':
            if idx < len(students):
                cell_value = "{0}({1})".format(students[idx].name, students[idx].id)
                c.value = cell_value
                idx += 1
            else:
                # clean X mark
                c.value = None
wb.save(new_seat_filename)
