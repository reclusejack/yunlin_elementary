#!/usr/bin/python3
# -*- coding: utf-8 -*-
import xlsxwriter
import random
import string
import sys
import names
import argparse
import importlib

importlib.reload

parser = argparse.ArgumentParser(
         description='please input teacher count and grade count')
parser.add_argument('-n','--number',
                    type=int,
                    default=100,
                    help='teacher count')
parser.add_argument('-g','--grade',
                    type=int,
                    default=6,
                    help='grade count')
args = parser.parse_args()

## show values ##
print ("Input number: %s" % args.number )

grade_list = ["一","二","三","四","五","六"]
workbook = xlsxwriter.Workbook('108鎮東國小導師名冊.xlsx')
teacher_sheet = workbook.add_worksheet('導師名冊')

teacher_list = []
grade = 0
x = 0
serial_number = 1
while (x < args.number):
    for g_n in range(1, int(int(args.number) / int(args.grade))):
        gen_teacher = []
        gen_teacher.append(serial_number)
        serial_number += 1
        gen_teacher.append(grade_list[grade])
        ran = random.randint(0, 1)
        if ran == 1:
                gen_teacher.append("女")
                gen_teacher.append(
                    names.get_first_name(gender='female')+str(x))
                gen_teacher.append("") #class
                gen_teacher.append("") #comment
                teacher_list.append(gen_teacher)
        else :
                gen_teacher.append("男")
                gen_teacher.append(
                    names.get_first_name(gender='male')+str(x))
                gen_teacher.append("") #class
                gen_teacher.append("") #comment
                teacher_list.append(gen_teacher)
    x += g_n
    grade += 1
    if grade == 6:
        break

#teacher title
gen_teacher = []
gen_teacher.append("序號")
gen_teacher.append("年級")
gen_teacher.append("性別")
gen_teacher.append("姓名")
gen_teacher.append("編訂班別")
gen_teacher.append("備註")
teacher_list.insert(0, gen_teacher)
gen_title = []
gen_title.append("108學年**國小導師編班名冊")
teacher_list.insert(0, gen_title)

col = 0
for row, data in enumerate(teacher_list):
        teacher_sheet.write_row(row, col, data)

#pdb.set_trace()
merge_format = workbook.add_format({
    'bold': 0,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter'})
teacher_sheet.merge_range('A1:F1',teacher_list[0][0], merge_format)

workbook.close()
