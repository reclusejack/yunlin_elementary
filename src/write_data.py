# -*- coding: utf-8 -*-
import xlsxwriter
import string
import re
import sys
import pprint


grade_list = ["一","二","三","四","五","六"]

def writefile(teacher_list, school_filename):
    #workbook = xlsxwriter.Workbook('108**國小導師名冊.xlsx')
    gen_title = []
    tmp = school_filename.replace(".xlsx", "")
    schoolname = re.search(r'.*\d+(.*)$', tmp).group(1)
    year = re.search(r'.*(\d+).*$', tmp).group(1)
    title_place = "雲林縣"
    title_year = "學年度國小導師編班名冊"
    gen_title = title_place + schoolname + year + title_year
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
    teacher_list.insert(0, gen_title)

    col = 0
    for row, data in enumerate(teacher_list):
            teacher_sheet.write_row(row, col, data)

    merge_format = workbook.add_format({
        'bold': 0,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter'})
    teacher_sheet.merge_range('A1:F1',teacher_list[0][0], merge_format)

    workbook.close()
    return
