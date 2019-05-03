# -*- coding: utf-8 -*-
import xlsxwriter
import string
import re
import sys
import datetime


grade_list = ["一","二","三","四","五","六"]

def writefile(teacher_list, school_filename):
    gen_title = []
    tmp = school_filename.replace(".xlsx", "")
    schoolname = re.search(r'.*(\d+)(.*)$', tmp).group(2)
    print(schoolname)
    year = str(int(datetime.date.today().year) - 1911)
    print(year)
    title_place = "雲林縣"
    append = "學年度"
    gen_title = title_place + year + append + schoolname
    school_filename = school_filename.replace("名冊.xlsx", "編定名冊.xlsx")
    school_filename = school_filename.replace("/", "\\")
    workbook = xlsxwriter.Workbook(school_filename)
    teacher_sheet = workbook.add_worksheet('導師名冊')

    #teacher title
    gen_teacher = []
    gen_teacher.append("序號")
    gen_teacher.append("年級")
    gen_teacher.append("性別")
    gen_teacher.append("姓名")
    gen_teacher.append("編訂班別")
    gen_teacher.append("備註")
    teacher_list.insert(0, gen_teacher)

    col = 0
    for row, data in enumerate(teacher_list):
            teacher_sheet.write_row(row + 1, col, data)

    teacher_list.insert(0, gen_title)
    merge_format = workbook.add_format({
        'bold': 0,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter'})
    teacher_sheet.merge_range('A1:G1',teacher_list[0], merge_format)

    workbook.close()
    return
