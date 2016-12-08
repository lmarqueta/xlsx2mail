#!/usr/bin/python
# -*- encoding utf-8 -*-
#
# Parses an Excel (xlsx) spreadsheet and creates a text list with urgent tasks
# (those with due date in a week or less) for a particular user, so it can be
# sent by email.
#
# Excel format should look like this:
# --------------------------------------------------------
# | Task | Status | Due | Owner | Effort | Pct | Comment |
#


import os
import sys
import argparse
import openpyxl
import warnings
import datetime


class Task:
    """Simple task definition"""
    def __init__(self, name, status, owner, due_date, comment):
        self.name = name.encode('utf-8')
        self.status = status
        if due_date:
            self.due_date = due_date.date()
        else:
            self.due_date = None
        self.owner = owner
        if comment:
            self.comment = comment.encode('utf-8')
        else:
            self.comment = None

    def show(self):
        print self.name

    def dtc(self):
        if self.due_date:
            now = datetime.date.today()
            delta = self.due_date - now
            return delta.days
        else:
            return None


def read_xlsx(xf):
    """Reads Excel xlsx file and creates a list of tasks
    :param xf: Path to Excel file
    """
    # We do not want (probably useless) warnings here
    warnings.simplefilter('ignore')

    task_list = []

    with open(xf) as fp:
        workbook = openpyxl.load_workbook(xf, read_only=True)
        sheet = workbook.get_sheet_by_name('Tasks')
        for row in range(2, sheet.max_row + 1):
            cell = 'B' + str(row)
            if sheet[cell].value:
                task_name = sheet[cell].value
                task_status = sheet['C'+str(row)].value
                task_due_date = sheet['D'+str(row)].value
                task_owner = sheet['E'+str(row)].value
                task_comment = sheet['I'+str(row)].value
                task = Task(task_name, task_status, task_owner, task_due_date, task_comment)
                task_list.append(task)

    return task_list


def main():
    # Command line options
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--filename', dest='filename', help='excel file (must be xlsx format', required='True')
    parser.add_argument('-u', '--username', dest='username', help='Username', required='True')
    args = parser.parse_args()

    filename = args.filename
    username = args.username

    # Make sure file exists
    if not os.path.isfile(filename):
        sys.stderr.write("Cannot read %s" % filename)
        sys.exit(1)

    # Parse Excel file
    tasks = read_xlsx(filename)

    # Get urgent tasks (due date in a week or less)
    urgent = []
    print "Tareas que deben terminar en una semana"
    print "======================================="
    for task in tasks:
        if task.owner == username and task.status == 'In progress' and task.dtc() and task.dtc() < 7:
            urgent.append([task.due_date, task.name, task.comment])

    urgent.sort(key=lambda due: due[0])
    for j in urgent:
        print j[0], j[1]
        if j[2]:
            print "          ", j[2]
    print

    # Other pending tasks
    print "Otras tareas pendientes"
    print "======================="
    other = []
    for task in tasks:
        if task.owner == username and not task.status == 'Finished' and task.dtc() >= 7:
            other.append([task.due_date, task.name, task.comment])

    other.sort(key=lambda due: due[0])
    for j in other:
        print j[0], j[1]
        if j[2]:
            print "          ", j[2]
    print

if __name__ == '__main__':
    main()
