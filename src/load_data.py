# -*- coding: utf-8 -*-
from openpyxl import load_workbook
import copy, re
import pprint
def load_data(school_filename):
    """load_data from excel file.

    :param school_filename: file name
    :param all_data: teacher/student information from file
    :param rule: same/diff notation table from excel file
    """
    wb = load_workbook(school_filename, read_only=True)
    ws_teacher = wb['導師名冊'] # ws is now an IterableWorksheet

    teacher_list = []
    name_ = '姓名'

    skip_row = 1 # skip first row
    for row in ws_teacher.rows:
        teacher_template = []
        #read each cell from xlsx
        for cell in row:
            if cell.value == name_:
                skip_row = 1
                continue
            if skip_row == 0:
                if (cell.value == None):
                    teacher_template.append(" ")
                else:
                    teacher_template.append(cell.value)
        if skip_row == 0:
            teacher_list.append(teacher_template)
        if skip_row == 1:
            skip_row = 0

    return teacher_list

