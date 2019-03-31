# -*- coding: utf-8 -*-
import argparse
import random
import copy
import string
import re
import datetime
#import win32com.client

#from openpyxl import load_workbook

th = 0
by = 1
gl = 2
sp = 3
same = 0
diff = 1
subtotal = []

# write outputfile

def writefile(group_data, school_filename, input_all, input_boy, input_girl):
    return
#    """writefile
#
#    :param group_data: all the data including teachers and students.
#    :param school_filename: school name for output
#    :param input_all: total class count
#    :param input_boy: boy class count
#    :param input_girl: girl class count
#    """
#    segregated = 0
#    class_all = 0
#    class_boy = 0
#    class_girl = 0
#    if  input_all == "0":
#        segregated = 1
#        class_boy = int(input_boy)
#        class_girl = int(input_girl)
#        class_all = class_boy + class_girl
#    else :
#        segregated = 0
#        class_all = int(input_all)
#        class_boy = class_all
#        class_girl = class_all
#
#    #Excel_output = win32com.client.gencache.EnsureDispatch('Excel.Application') # Excel = win32com.client.Dispatch('Excel.Application')
#
#    #Excel_input.Visible = 1
#
#    Excel_output.Visible = 1
#    win32c = win32com.client.constants
#
#    wb_out = Excel_output.Workbooks.Add()
#
#    wb_out.Worksheets.Add(None,wb_out.Worksheets(1),None,None)
#    wb_out.Worksheets.Add(None,wb_out.Worksheets(2),None,None)
#    wb_out.Worksheets.Add(None,wb_out.Worksheets(3),None,None)
#    Sheet1 = wb_out.Worksheets(1)
#    Sheet1.Name = "導師名冊"
#    Sheet2 = wb_out.Worksheets(2)
#    Sheet2.Name = "男生名冊"
#    Sheet3 = wb_out.Worksheets(3)
#    Sheet3.Name = "女生名冊"
#    Sheet4 = wb_out.Worksheets(4)
#    Sheet4.Name = "特殊生名冊"
#
#
#    #teacher title
#    gen_teacher = []
#    gen_teacher.append("導師序號")
#    gen_teacher.append("導師姓名")
#    gen_teacher.append("導師性別")
#    gen_teacher.append("導師班別")
#    gen_teacher.append("導師備註")
#    group_data[th].insert(0, gen_teacher)
#
#    tmp = school_filename.replace(".xlsx", "")
#    schoolname = re.search(r'.*\d+(.*)$', tmp).group(1)
#    title_place = "雲林縣"
#    year = str(int(datetime.date.today().year) - 1911)
#    title_year = "學年度新生常態編班"
#    title_th = "導師名冊"
#    title_st = "學生名冊"
#    title_by = "男生名冊"
#    title_gl = "女生名冊"
#    title_sp = "特殊生名冊"
#    title_result = "班級一覽表"
#    title = title_place + schoolname + year + title_year + title_th
#    gen_title = []
#    gen_title.append(title)
#    group_data[th].insert(0, gen_title)
#
## write teacher tab
#    row = 0
#    col = 0
#    for row, data in enumerate(group_data[th]):
#        for values in data:
#            if col >= 5:
#                continue
#            Sheet1.Cells(row+1, col+1).Value = values
#            col = col + 1
#        col = 0
#    cl1 = Sheet1.Cells(1, 1)
#    cl2 = Sheet1.Cells(1, len(group_data[th][1]))
#    selection = Sheet1.Range(cl1,cl2)
#    selection.MergeCells = True
#    cl2 = Sheet1.Cells(len(group_data[th]), len(group_data[th][1]))
#    Sheet1.Range(Sheet1.Cells(1,1), cl2).VerticalAlignment = 2
#    Sheet1.Range(Sheet1.Cells(1,1), cl2).HorizontalAlignment = 3
#    for border_id in range(7,13):
#        Sheet1.Range(Sheet1.Cells(1,1), cl2).Borders(border_id).Weight = 2
#        Sheet1.Range(Sheet1.Cells(1,1), cl2).Borders(border_id).LineStyle = 1
#    Sheet1.Columns(5).WrapText = True
#    Sheet1.Columns(5).ColumnWidth = 30
#
## write teacher name reference
#    n = 1
#    for data in gen_teacher :
#        cl1 = Sheet1.Cells(3, n)
#        cl2 = Sheet1.Cells(len(group_data[th]), n)
#        selection = Sheet1.Range(cl1,cl2)
#        wb_out.Names.Add(Name=gen_teacher[n-1], RefersTo=selection)
#        n = n + 1
#
#    #student title
#    gen_student = []
#    gen_student.append("序號")
#    gen_student.append("姓名")
#    gen_student.append("性別")
#    gen_student.append("成績")
#    gen_student.append("編訂班別")
#    gen_student.append("同班")
#    gen_student.append("不同班")
#    gen_student.append("同班/不同班 序號")
#    gen_student.append("備註")
##    group_data[st].insert(0, gen_student)
#    group_data[by].insert(0, gen_student)
#    group_data[gl].insert(0, gen_student)
#    gen_special_student = []
#    gen_special_student.append("序號")
#    gen_special_student.append("姓名")
#    gen_special_student.append("性別")
#    gen_special_student.append("成績")
#    gen_special_student.append("編訂班別")
#    gen_special_student.append("備註")
#    group_data[sp].insert(0, gen_special_student)
#
#    gen_title = []
#    title = title_place + schoolname + year + title_year + title_by
#    gen_title.append(title)
#    group_data[by].insert(0, gen_title)
#    gen_title = []
#    title = title_place + schoolname + year + title_year + title_gl
#    gen_title.append(title)
#    group_data[gl].insert(0, gen_title)
#    gen_title = []
#    title = title_place + schoolname + year + title_year + title_sp
#    gen_title.append(title)
#    group_data[sp].insert(0, gen_title)
#    gen_title = []
#    title = title_place + schoolname + year + title_year + title_result
#    gen_title.append(title)
## write boy tab
#    row = 0
#    col = 0
#    for row, data in enumerate(group_data[by]):
#        for values in data:
#            if col >= 9:
#                continue
#            if values == "":
#                values = " "
#            Sheet2.Cells(row+1, col+1).Value = values
#            col = col + 1
#        col = 0
#    cl1 = Sheet2.Cells(1, 1)
#    cl2 = Sheet2.Cells(1, len(group_data[by][1]))
#    selection = Sheet2.Range(cl1,cl2)
#    selection.MergeCells = True
#    cl2 = Sheet2.Cells(len(group_data[by]), len(group_data[by][1]))
#    Sheet2.Range(Sheet2.Cells(1,1), cl2).VerticalAlignment = 2
#    Sheet2.Range(Sheet2.Cells(1,1), cl2).HorizontalAlignment = 3
#    for border_id in range(7,13):
#        Sheet2.Range(Sheet2.Cells(1,1), cl2).Borders(border_id).Weight = 2
#        Sheet2.Range(Sheet2.Cells(1,1), cl2).Borders(border_id).LineStyle = 1
#    Sheet2.Columns(9).WrapText = True
#    Sheet2.Columns(1).ColumnWidth = 7
#    Sheet2.Columns(2).ColumnWidth = 9
#    Sheet2.Columns(3).ColumnWidth = 5
#    Sheet2.Columns(4).ColumnWidth = 8
#    Sheet2.Columns(5).ColumnWidth = 8
#    Sheet2.Columns(6).ColumnWidth = 5.5
#    Sheet2.Columns(7).ColumnWidth = 6.5
#    Sheet2.Columns(8).ColumnWidth = 6.5
#    Sheet2.Columns(9).ColumnWidth = 15
#    Sheet2.PageSetup.PrintTitleRows = "$2:$2"
## write girl tab
#    row = 0
#    col = 0
#    for row, data in enumerate(group_data[gl]):
#        for values in data:
#            if col >= 9:
#                continue
#            if values == "":
#                values = " "
#            Sheet3.Cells(row+1, col+1).Value = values
#            col = col + 1
#        col = 0
#    cl1 = Sheet3.Cells(1, 1)
#    cl2 = Sheet3.Cells(1, len(group_data[gl][1]))
#    selection = Sheet3.Range(cl1,cl2)
#    selection.MergeCells = True
#    cl2 = Sheet3.Cells(len(group_data[gl]), len(group_data[gl][1]))
#    Sheet3.Range(Sheet3.Cells(1,1), cl2).VerticalAlignment = 2
#    Sheet3.Range(Sheet3.Cells(1,1), cl2).HorizontalAlignment = 3
#    for border_id in range(7,13):
#        Sheet3.Range(Sheet3.Cells(1,1), cl2).Borders(border_id).Weight = 2
#        Sheet3.Range(Sheet3.Cells(1,1), cl2).Borders(border_id).LineStyle = 1
#    Sheet3.Columns(9).WrapText = True
#    Sheet3.Columns(1).ColumnWidth = 7
#    Sheet3.Columns(2).ColumnWidth = 9
#    Sheet3.Columns(3).ColumnWidth = 5
#    Sheet3.Columns(4).ColumnWidth = 8
#    Sheet3.Columns(5).ColumnWidth = 8
#    Sheet3.Columns(6).ColumnWidth = 5.5
#    Sheet3.Columns(7).ColumnWidth = 6.5
#    Sheet3.Columns(8).ColumnWidth = 6.5
#    Sheet3.Columns(9).ColumnWidth = 15
#    Sheet3.PageSetup.PrintTitleRows = "$2:$2"
## write special tab
#    row = 0
#    col = 0
#    for row, data in enumerate(group_data[sp]):
#        for values in data:
#            if col >= 6:
#                continue
#            if values == "":
#                values = " "
#            Sheet4.Cells(row+1, col+1).Value = values
#            col = col + 1
#        col = 0
#    cl1 = Sheet4.Cells(1, 1)
#    cl2 = Sheet4.Cells(1, len(group_data[sp][1]))
#    Sheet4.Columns(6).WrapText = True
#    Sheet4.Columns(1).ColumnWidth = 7
#    Sheet4.Columns(2).ColumnWidth = 9
#    Sheet4.Columns(3).ColumnWidth = 5
#    Sheet4.Columns(4).ColumnWidth = 8
#    Sheet4.Columns(5).ColumnWidth = 8
#    Sheet4.Columns(6).ColumnWidth = 40
#    selection = Sheet4.Range(cl1,cl2)
#    selection.MergeCells = True
#    cl2 = Sheet4.Cells(len(group_data[sp]), len(group_data[sp][1]))
#    Sheet4.Range(Sheet4.Cells(1,1), cl2).VerticalAlignment = 2
#    Sheet4.Range(Sheet4.Cells(1,1), cl2).HorizontalAlignment = 3
#    for border_id in range(7,13):
#        Sheet4.Range(Sheet4.Cells(1,1), cl2).Borders(border_id).Weight = 2
#        Sheet4.Range(Sheet4.Cells(1,1), cl2).Borders(border_id).LineStyle = 1
## write special student name reference
#    n = 1
#    for data in gen_special_student :
#        cl1 = Sheet4.Cells(3, n)
#        cl2 = Sheet4.Cells(len(group_data[sp]), n)
#        selection = Sheet4.Range(cl1,cl2)
#        wb_out.Names.Add(Name=gen_special_student[n-1], RefersTo=selection)
#        n = n + 1
#
#
#    for current_tab in range(1, class_boy+1):
#
#       #boy in each tab
#        wb_out.Worksheets.Add(None,wb_out.Worksheets(current_tab + 3),None,None)
#        Sheet_class_1 = wb_out.Worksheets(current_tab + 4)
#
#        Sheet_class_1.Select()
#
#        class_name = 100 + current_tab
#        Sheet_class_1.Name = str(class_name)
#
## print teacher
#        th_start_col = 0
#        col = 1
#        row = 1
#        for data in gen_teacher :
#            Sheet_class_1.Cells(row, col + th_start_col).Value = data
#            col = col + 1
#
#        col = 1
#        for data in gen_teacher:
#            formula_th1 = "=IFERROR(INDEX(INDIRECT(導師名冊!"
#            formula_th2 = string.ascii_uppercase[col - 1]
#            formula_th3 = "$2),SMALL(IF(導師班別="
#            formula_th4 = ",ROW(導師班別),FALSE),ROW("
#            formula_th5 = "))-2,1),\"\")"
#            formula_th = formula_th1 + formula_th2 + formula_th3+ str(current_tab) + formula_th4
#            formula_th = formula_th + "1:1" + formula_th5
#
#            cl1 = Sheet_class_1.Cells(row + 1, col + th_start_col)
#            cl2 = Sheet_class_1.Cells(row + 1, col + th_start_col)
#            Sheet_class_1.Range(cl1,cl2).FormulaArray = formula_th
#            col = col + 1
#
#        Sheet2.Select()
#        cl1 = Sheet2.Cells(2,1)
#        cl2 = Sheet2.Cells(len(group_data[by]),len(group_data[by][2]))
#        PivotSourceRange = Sheet2.Range(cl1,cl2)
#
#        PivotSourceRange.Select()
#
#
#        cl3=Sheet_class_1.Cells(5,1)
#        PivotTargetRange = Sheet_class_1.Range(cl3,cl3)
#        PivotTableName = 'ReportPivotTable'
#
#        PivotCache = wb_out.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)
#
#        PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)
#
#
#        PivotTable.PivotFields('編訂班別').Orientation = win32c.xlPageField
#        PivotTable.PivotFields('編訂班別').Position = 1
#        PivotTable.PivotFields('編訂班別').CurrentPage = str(current_tab)
#        PivotTable.PivotFields('序號').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('序號').Position = 1
#        PivotTable.PivotFields('序號').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('姓名').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('姓名').Position = 2
#        PivotTable.PivotFields('姓名').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('性別').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('性別').Position = 3
#        PivotTable.PivotFields('性別').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('同班').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('同班').Position = 4
#        PivotTable.PivotFields('同班').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('不同班').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('不同班').Position = 5
#        PivotTable.PivotFields('不同班').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('同班/不同班 序號').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('同班/不同班 序號').Position = 6
#        PivotTable.PivotFields('同班/不同班 序號').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('備註').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('備註').Position = 7
#        PivotTable.PivotFields('備註').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        DataField = PivotTable.AddDataField(PivotTable.PivotFields('成績'))
#        DataField.Function = 2 # 2 = Avarage
#        DataField.Name = '成績 '
#        PivotTable.PivotFields('序號').AutoSort(2, "成績 ")
#        PivotTable.RowAxisLayout(1)
#        PivotTable.ShowDrillIndicators = 0
#
#
#        if segregated == 1:
#            sp_start_col = 0
#            row = int(len(group_data[by])/int(input_boy)) + 12
#            sp_start_row = row
#            col = 1
#
#            for data in gen_special_student:
#                Sheet_class_1.Cells(row, col + sp_start_col).Value = data
#                col = col + 1
#
#            col = 1
#            for offset in range(1,5):
#                col = 1
#                for data in gen_special_student:
#                    formula1 = "=IFERROR(INDEX(INDIRECT(特殊生名冊!"
#                    formula2 = string.ascii_uppercase[col - 1]
#                    formula3 = "$2),SMALL(IF(編訂班別="
#                    formula4 = ",ROW(編訂班別),FALSE),ROW("
#                    formula5 = "))-2,1),\"\")"
#                    formula = formula1 + formula2 + formula3+ str(current_tab) + formula4
#                    formula = formula + str(offset) + ":" + str(offset) + formula5
#
#                    cl1 = Sheet_class_1.Cells(offset + sp_start_row, col +sp_start_col)
#                    cl2 = Sheet_class_1.Cells(offset + sp_start_row, col +sp_start_col)
#                    Sheet_class_1.Range(cl1,cl2).FormulaArray = formula
#                    col = col + 1
#
#            Sheet_class_1.Columns.AutoFit()
#            Sheet_class_1.Columns(1).ColumnWidth = 7.5
#            Sheet_class_1.Columns(2).ColumnWidth = 7.5
#            Sheet_class_1.Columns(3).ColumnWidth = 7.5
#            Sheet_class_1.Columns(4).ColumnWidth = 7.5
#            Sheet_class_1.Columns(5).ColumnWidth = 7.75
#            Sheet_class_1.Columns(6).ColumnWidth = 4.5
#            Sheet_class_1.Columns(7).ColumnWidth = 24
#            Sheet_class_1.Columns(7).WrapText = True
#            Sheet_class_1.Columns(8).ColumnWidth = 8
#            Sheet_class_1.Range(Sheet_class_1.Cells(1,1), Sheet_class_1.Cells(100,10)).Font.Size = 11
#            Sheet_class_1.PageSetup.BottomMargin = 1
#
#
#    for current_tab in range(1,class_girl+1):
#
#        #girl in each tab
#
#        Sheet3.Select()
#        cl1 = Sheet3.Cells(2,1)
#        cl2 = Sheet3.Cells(len(group_data[gl]),9)
#        PivotSourceRange = Sheet3.Range(cl1,cl2)
#        PivotSourceRange.Select()
#
#        if segregated == 0:
#            gl_start_row = int(len(group_data[by])/int(input_all)) + 12
#            Sheet_class_1 = wb_out.Worksheets(current_tab + 4)
#            Sheet_class_1.Select()
#            cl3=Sheet_class_1.Cells(gl_start_row,1)
#        else :
#            wb_out.Worksheets.Add(None,wb_out.Worksheets(class_boy + current_tab + 3),None,None)
#            Sheet_class_1 = wb_out.Worksheets(class_boy + current_tab + 4)
#            Sheet_class_1.Select()
#            class_name = 100 + current_tab + class_boy
#            Sheet_class_1.Name = str(class_name)
#
#            th_start_col = 0
#            col = 1
#            row = 1
#            for data in gen_teacher :
#                Sheet_class_1.Cells(row, col + th_start_col).Value = data
#                col = col + 1
#
#            col = 1
#            for data in gen_teacher:
#                formula_th1 = "=IFERROR(INDEX(INDIRECT(導師名冊!"
#                formula_th2 = string.ascii_uppercase[col - 1]
#                formula_th3 = "$2),SMALL(IF(導師班別="
#                formula_th4 = ",ROW(導師班別),FALSE),ROW("
#                formula_th5 = "))-2,1),\"\")"
#                formula_th = formula_th1 + formula_th2 + formula_th3+ str(current_tab + class_boy) + formula_th4
#                formula_th = formula_th + "1:1" + formula_th5
#
#                cl1 = Sheet_class_1.Cells(row + 1, col + th_start_col)
#                cl2 = Sheet_class_1.Cells(row + 1, col + th_start_col)
#                Sheet_class_1.Range(cl1,cl2).FormulaArray = formula_th
#                col = col + 1
#
#            cl3=Sheet_class_1.Cells(5,1)
#
#        PivotTargetRange = Sheet_class_1.Range(cl3,cl3)
#        PivotTableName = 'ReportPivotTable_2'
#
#        PivotCache = wb_out.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)
#
#        PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)
#
#
#        PivotTable.PivotFields('編訂班別').Orientation = win32c.xlPageField
#        PivotTable.PivotFields('編訂班別').Position = 1
#        PivotTable.PivotFields('編訂班別').CurrentPage = str(current_tab)
#        PivotTable.PivotFields('序號').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('序號').Position = 1
#        PivotTable.PivotFields('序號').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('姓名').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('姓名').Position = 2
#        PivotTable.PivotFields('姓名').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('性別').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('性別').Position = 3
#        PivotTable.PivotFields('性別').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('同班').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('同班').Position = 4
#        PivotTable.PivotFields('同班').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('不同班').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('不同班').Position = 5
#        PivotTable.PivotFields('不同班').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('同班/不同班 序號').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('同班/不同班 序號').Position = 6
#        PivotTable.PivotFields('同班/不同班 序號').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        PivotTable.PivotFields('備註').Orientation = win32c.xlRowField
#        PivotTable.PivotFields('備註').Position = 7
#        PivotTable.PivotFields('備註').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
#        DataField = PivotTable.AddDataField(PivotTable.PivotFields('成績'))
#        DataField.Function = 2 # 2 = Avarage
#        DataField.Name = '成績 '
#        PivotTable.PivotFields('序號').AutoSort(2, "成績 ")
#        PivotTable.RowAxisLayout(1)
#        PivotTable.ShowDrillIndicators = 0
#
#
#        if segregated == 0:
#            sp_start_col = 0
#            row = gl_start_row + int(len(group_data[gl])/int(input_all)) + 3
#            sp_start_row = row
#            col = 1
#        else :
#            sp_start_col = 0
#            row = int(len(group_data[gl])/int(input_girl)) + 12
#            sp_start_row = row
#            col = 1
#
#        for data in gen_special_student:
#            Sheet_class_1.Cells(row, col + sp_start_col).Value = data
#            col = col + 1
#
#        col = 1
#        for offset in range(1,5):
#            col = 1
#            for data in gen_special_student:
#                formula1 = "=IFERROR(INDEX(INDIRECT(特殊生名冊!"
#                formula2 = string.ascii_uppercase[col - 1]
#                formula3 = "$2),SMALL(IF(編訂班別="
#                formula4 = ",ROW(編訂班別),FALSE),ROW("
#                formula5 = "))-2,1),\"\")"
#                if segregated == 0:
#                    formula = formula1 + formula2 + formula3+ str(current_tab) + formula4
#                else:
#                    formula = formula1 + formula2 + formula3+ str(current_tab + class_boy) + formula4
#                formula = formula + str(offset) + ":" + str(offset) + formula5
#
#                cl1 = Sheet_class_1.Cells(offset + sp_start_row, col +sp_start_col)
#                cl2 = Sheet_class_1.Cells(offset + sp_start_row, col +sp_start_col)
#                Sheet_class_1.Range(cl1,cl2).FormulaArray = formula
#                col = col + 1
#        Sheet_class_1.Columns.AutoFit()
#        Sheet_class_1.Columns(1).ColumnWidth = 7.5
#        Sheet_class_1.Columns(2).ColumnWidth = 7.5
#        Sheet_class_1.Columns(3).ColumnWidth = 7.5
#        Sheet_class_1.Columns(4).ColumnWidth = 7.5
#        Sheet_class_1.Columns(5).ColumnWidth = 7.75
#        Sheet_class_1.Columns(6).ColumnWidth = 4.5
#        Sheet_class_1.Columns(7).ColumnWidth = 24
#        Sheet_class_1.Columns(7).WrapText = True
#        Sheet_class_1.Columns(8).ColumnWidth = 8
#        Sheet_class_1.Range(Sheet_class_1.Cells(1,1), Sheet_class_1.Cells(100,10)).Font.Size = 11
#        Sheet_class_1.PageSetup.BottomMargin = 1
#
##### update grouping status #############
#
#    wb_out.Worksheets.Add(None,wb_out.Worksheets(class_all + 4),None,None)
#    Sheet_status = wb_out.Worksheets(class_all + 5)
#    Sheet_status.Select()
#    result_name = "編班結果"
#    Sheet_status.Name = str(result_name)
#
#    if segregated == 0:
#        for class_result in range(1,3):
#            if class_result == 1:
#                Sheet2.Select()
#                cl1 = Sheet2.Cells(2,1)
#                cl2 = Sheet2.Cells(len(group_data[by]),len(group_data[by][2]))
#                PivotSourceRange = Sheet2.Range(cl1,cl2)
#
#                PivotSourceRange.Select()
#
#                cl3=Sheet_status.Cells(1,1)
#                cl4=Sheet_status.Cells(4,1)
#                PivotTableName = 'ReportPivotTable_by'
#                PivotTableName_count = 'ReportPivotTable_by_count'
#            else:
#                Sheet3.Select()
#                cl1 = Sheet3.Cells(2,1)
#                #cl2 = Sheet3.Cells(len(group_data[gl]),len(group_data[gl][2]))
#                cl2 = Sheet3.Cells(len(group_data[gl]),9)
#                PivotSourceRange = Sheet3.Range(cl1,cl2)
#
#                PivotSourceRange.Select()
#
#                cl3=Sheet_status.Cells(7,1)
#                cl4=Sheet_status.Cells(10,1)
#                PivotTableName = 'ReportPivotTable_gl'
#                PivotTableName_count = 'ReportPivotTable_gl_count'
#
#            PivotTargetRange = Sheet_status.Range(cl3,cl3)
#            PivotCache = wb_out.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)
#
#            PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)
#
#            PivotTable.PivotFields('編訂班別').Orientation = win32c.xlColumnField
#            PivotTable.PivotFields('編訂班別').Position = 1
#            DataField = PivotTable.AddDataField(PivotTable.PivotFields('成績'))
#            DataField.Function = 2 # 2 = Avarage
#            DataField.Name = '平均成績'
#
#            PivotTargetRange_count = Sheet_status.Range(cl4,cl4)
#            PivotCache_count = wb_out.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)
#
#            PivotTable_count = PivotCache_count.CreatePivotTable(TableDestination=PivotTargetRange_count, TableName=PivotTableName_count, DefaultVersion=win32c.xlPivotTableVersion14)
#
#            PivotTable_count.PivotFields('編訂班別').Orientation = win32c.xlColumnField
#            PivotTable_count.PivotFields('編訂班別').Position = 1
#            DataField = PivotTable_count.AddDataField(PivotTable_count.PivotFields('序號'))
#            DataField.Name = '人數'
#
## cacluate average score
#        avg_row = 15
#        avg_colmn = 1
#        avg_out_row = 16
#        avg_out_colmn = 1
#        Sheet_status.Columns(1).ColumnWidth = 10
#        Sheet_status.Columns(2).ColumnWidth = 10
#        Sheet_status.Columns(3).ColumnWidth = 10
#        Sheet_status.Columns(4).ColumnWidth = 10
#        Sheet_status.Columns(5).ColumnWidth = 10
#        Sheet_status.Columns(6).ColumnWidth = 12
#        Sheet_status.Columns(7).ColumnWidth = 12
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn).Value = "班級"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 1).Value = "平均成績"
#        for class_num in range(0, class_all):
#            class_name = 100 + class_num + 1
#            Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn).Value = str(class_name)
#
#        for class_num in range(0, class_all):
#            f_col = string.ascii_uppercase[avg_colmn + class_num]
#            f_avg1 = "=" + "(" + f_col + "3" + "*" + f_col + "6" + "+"
#            f_avg2 = f_col + "9" + "*" + f_col + "12" + ")" + "/"
#            f_avg3 = "(" + f_col + "6" + "+" + f_col + "12" + ")"
#
#            formula_avg = f_avg1 + f_avg2 + f_avg3
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 1)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 1)
#            Sheet_status.Range(sel1,sel2).Formula = formula_avg
## class_average
#        f_col = string.ascii_uppercase[avg_colmn]
#        f_avg = "=average(" + f_col + str(avg_out_row) + ":" + f_col + str(avg_out_row + class_all -1) + ")"
#
#        formula_avg = f_avg
#        sel1 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + 1)
#        sel2 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + 1)
#        Sheet_status.Range(sel1,sel2).Formula = formula_avg
## stdev
#        Sheet_status.Cells(avg_out_row + class_all , avg_out_colmn).Value = "班級平均"
#        Sheet_status.Cells(avg_out_row + class_all + 1, avg_out_colmn).Value = "標準差"
#        f_col = string.ascii_uppercase[avg_colmn]
#        f_avg1 = "=stdev(" + f_col + str(avg_out_row) + ":" + f_col + str(avg_out_row + class_all -1) + ")"
#
#        formula_avg = f_avg1
#        sel1 = Sheet_status.Cells(avg_out_row + class_all + 1, avg_out_colmn + 1)
#        sel2 = Sheet_status.Cells(avg_out_row + class_all + 1, avg_out_colmn + 1)
#        Sheet_status.Range(sel1,sel2).Formula = formula_avg
#
##student count
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 2).Value = "男生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 3).Value = "女生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 4).Value = "學生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 5).Value = "特殊生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 6).Value = "班級總人數"
#
#        for class_num in range(0, class_all):
#            f_col = string.ascii_uppercase[avg_colmn + class_num]
#            formula_avg = "=" + f_col + str(6)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+2)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+2)
#            Sheet_status.Range(sel1,sel2).Formula = formula_avg
#
#        for class_num in range(0, class_all):
#            f_col = string.ascii_uppercase[avg_colmn + class_num]
#            formula_avg = "=" + f_col + str(12)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+3)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+3)
#            Sheet_status.Range(sel1,sel2).Formula = formula_avg
#
#        for st_sum in range(2,7):
#            f_col = string.ascii_uppercase[avg_colmn + st_sum - 1]
#            f_sum = "=sum(" + f_col + str(avg_out_row) + ":" + f_col + str(avg_out_row + class_all -1) + ")"
#            sel1 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + st_sum)
#            sel2 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + st_sum)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
#
#        for class_num in range(0, class_all):
#            f_sum = "=" + "C" + str(avg_out_row + class_num) + "+" + "D" + str(avg_out_row + class_num)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 4)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 4)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
## Count special studnet count in each class
#        for class_num in range(0, class_all):
#            f_sum = "=COUNTIF(特殊生名冊!E3:E" + str(len(group_data[sp])) + ",\""+str(class_num + 1)+"\")"
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 5)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 5)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
## Count total student count in each class
#        for class_num in range(0, class_all):
#            f_sum = "=" + "E" + str(avg_out_row + class_num) + "+" + "F" + str(avg_out_row + class_num)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 6)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 6)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
## hide rows
#        Sheet_status.Cells(avg_row - 1, 1).Value = gen_title
#        sel1 = Sheet_status.Cells(avg_row - 1,1)
#        sel2 = Sheet_status.Cells(avg_row - 1,7)
#        Sheet_status.Range(sel1,sel2).MergeCells = True
#        Sheet_status.Range(sel1,sel2).HorizontalAlignment = 3
#        sel2 = Sheet_status.Cells(avg_out_row + class_all + 1, 7)
#        for border_id in range(7,13):
#            Sheet_status.Range(sel1,sel2).Borders(border_id).Weight = 2
#            Sheet_status.Range(sel1,sel2).Borders(border_id).LineStyle = 1
#        Sheet_status.Rows("1:13").Hidden = True
#
#    else:
#        for class_result in range(1,3):
#            if class_result == 1:
#                if (class_boy == 0):
#                    continue;
#                Sheet2.Select()
#                cl1 = Sheet2.Cells(2,1)
#                cl2 = Sheet2.Cells(len(group_data[by]),len(group_data[by][2]))
#                PivotSourceRange = Sheet2.Range(cl1,cl2)
#
#                PivotSourceRange.Select()
#
#                cl3=Sheet_status.Cells(1,1)
#                cl4=Sheet_status.Cells(4,1)
#                PivotTableName = 'ReportPivotTable_by'
#                PivotTableName_count = 'ReportPivotTable_by_count'
#            else:
#                if (class_girl == 0):
#                    continue;
#                Sheet3.Select()
#                cl1 = Sheet3.Cells(2,1)
#                cl2 = Sheet3.Cells(len(group_data[gl]),9)
#                PivotSourceRange = Sheet3.Range(cl1,cl2)
#
#                PivotSourceRange.Select()
#
#                cl3=Sheet_status.Cells(7,1)
#                cl4=Sheet_status.Cells(10,1)
#                PivotTableName = 'ReportPivotTable_gl'
#                PivotTableName_count = 'ReportPivotTable_gl_count'
#
#            PivotTargetRange = Sheet_status.Range(cl3,cl3)
#            PivotCache = wb_out.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)
#
#            PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)
#
#            PivotTable.PivotFields('編訂班別').Orientation = win32c.xlColumnField
#            PivotTable.PivotFields('編訂班別').Position = 1
#            DataField = PivotTable.AddDataField(PivotTable.PivotFields('成績'))
#            DataField.Function = 2 # 2 = Avarage
#            DataField.Name = '平均成績'
#
#            PivotTargetRange_count = Sheet_status.Range(cl4,cl4)
#            PivotCache_count = wb_out.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)
#
#            PivotTable_count = PivotCache_count.CreatePivotTable(TableDestination=PivotTargetRange_count, TableName=PivotTableName_count, DefaultVersion=win32c.xlPivotTableVersion14)
#
#            PivotTable_count.PivotFields('編訂班別').Orientation = win32c.xlColumnField
#            PivotTable_count.PivotFields('編訂班別').Position = 1
#            DataField = PivotTable_count.AddDataField(PivotTable_count.PivotFields('序號'))
#            DataField.Name = '人數'
#
## cacluate average score
#        avg_row = 15
#        avg_colmn = 1
#        avg_out_row = 16
#        avg_out_colmn = 1
#        Sheet_status.Columns(1).ColumnWidth = 10
#        Sheet_status.Columns(2).ColumnWidth = 10
#        Sheet_status.Columns(3).ColumnWidth = 10
#        Sheet_status.Columns(4).ColumnWidth = 10
#        Sheet_status.Columns(5).ColumnWidth = 10
#        Sheet_status.Columns(6).ColumnWidth = 12
#        Sheet_status.Columns(7).ColumnWidth = 12
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn).Value = "班級"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 1).Value = "平均成績"
#        for class_num in range(0, class_all):
#            class_name = 100 + class_num + 1
#            Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn).Value = str(class_name)
#
#        for class_num in range(0, class_boy):
#            f_col = string.ascii_uppercase[avg_colmn + class_num]
#            formula_avg = "=" + f_col + str(3)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+1)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+1)
#            Sheet_status.Range(sel1,sel2).Formula = formula_avg
#
#        for class_num in range(0, class_girl):
#            f_col = string.ascii_uppercase[avg_colmn + class_num]
#            formula_avg = "=" + f_col + str(9)
#            sel1 = Sheet_status.Cells(avg_out_row + class_boy + class_num, avg_out_colmn+1)
#            sel2 = Sheet_status.Cells(avg_out_row + class_boy + class_num, avg_out_colmn+1)
#            Sheet_status.Range(sel1,sel2).Formula = formula_avg
##student count
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 2).Value = "男生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 3).Value = "女生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 4).Value = "學生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 5).Value = "特殊生人數"
#        Sheet_status.Cells(avg_out_row - 1, avg_out_colmn + 6).Value = "班級總人數"
#        for class_num in range(0, class_boy):
#            f_col = string.ascii_uppercase[avg_colmn + class_num]
#            formula_avg = "=" + f_col + str(6)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+2)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn+2)
#            Sheet_status.Range(sel1,sel2).Formula = formula_avg
#
#        for class_num in range(0, class_girl):
#            f_col = string.ascii_uppercase[avg_colmn + class_num]
#            formula_avg = "=" + f_col + str(12)
#            sel1 = Sheet_status.Cells(avg_out_row + class_boy + class_num, avg_out_colmn+3)
#            sel2 = Sheet_status.Cells(avg_out_row + class_boy + class_num, avg_out_colmn+3)
#            Sheet_status.Range(sel1,sel2).Formula = formula_avg
#
#        for st_sum in range(2,7):
#            f_col = string.ascii_uppercase[avg_colmn + st_sum - 1]
#            f_sum = "=sum(" + f_col + str(avg_out_row) + ":" + f_col + str(avg_out_row + class_all -1) + ")"
#            sel1 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + st_sum)
#            sel2 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + st_sum)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
#
#        for class_num in range(0, class_all):
#            f_sum = "=" + "C" + str(avg_out_row + class_num) + "+" + "D" + str(avg_out_row + class_num)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 4)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 4)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
#
## Count special studnet count in each class
#        for class_num in range(0, class_all):
#            f_sum = "=COUNTIF(特殊生名冊!E3:E" + str(len(group_data[sp])) + ",\""+str(class_num + 1)+"\")"
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 5)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 5)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
## Count total student count in each class
#        for class_num in range(0, class_all):
#            f_sum = "=" + "E" + str(avg_out_row + class_num) + "+" + "F" + str(avg_out_row + class_num)
#            sel1 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 6)
#            sel2 = Sheet_status.Cells(avg_out_row + class_num, avg_out_colmn + 6)
#            Sheet_status.Range(sel1,sel2).Formula = f_sum
#
## class_average
#        Sheet_status.Cells(avg_out_row + class_all , avg_out_colmn).Value = "班級平均"
#        Sheet_status.Cells(avg_out_row + class_all + 1, avg_out_colmn).Value = "標準差"
#        f_col = string.ascii_uppercase[avg_colmn]
#        f_avg = "=average(" + f_col + str(avg_out_row) + ":" + f_col + str(avg_out_row + class_all -1) + ")"
#
#        formula_avg = f_avg
#        sel1 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + 1)
#        sel2 = Sheet_status.Cells(avg_out_row + class_all, avg_out_colmn + 1)
#        Sheet_status.Range(sel1,sel2).Formula = formula_avg
## stdev
#        f_col = string.ascii_uppercase[avg_colmn]
#        f_stdev = "=stdev(" + f_col + str(avg_out_row) + ":" + f_col + str(avg_out_row + class_all -1) + ")"
#
#        formula_stdev = f_stdev
#        sel1 = Sheet_status.Cells(avg_out_row + class_all + 1, avg_out_colmn + 1)
#        sel2 = Sheet_status.Cells(avg_out_row + class_all + 1, avg_out_colmn + 1)
#        Sheet_status.Range(sel1,sel2).Formula = formula_stdev
## hide rows
#        Sheet_status.Cells(avg_row - 1, 1).Value = gen_title
#        sel1 = Sheet_status.Cells(avg_row - 1,1)
#        sel2 = Sheet_status.Cells(avg_row - 1,7)
#        Sheet_status.Range(sel1,sel2).MergeCells = True
#        Sheet_status.Range(sel1,sel2).HorizontalAlignment = 3
#        sel2 = Sheet_status.Cells(avg_out_row + class_all + 1, 7)
#        for border_id in range(7,13):
#            Sheet_status.Range(sel1,sel2).Borders(border_id).Weight = 2
#            Sheet_status.Range(sel1,sel2).Borders(border_id).LineStyle = 1
#        Sheet_status.Rows("1:13").Hidden = True
##### write file #############
#
#    Sheet1.Select()
#    school_filename = school_filename.replace(".xlsx", "編班名冊.xlsx")
#    school_filename = school_filename.replace("/", "\\")
##    print ("write filename update: %s" %school_filename)
#
#    for n in range(1, class_all + 6):
#        Sheet_active = wb_out.Worksheets(n)
#        Sheet_active.Select()
#        Sheet_active.PageSetup.CenterFooter = "第&P頁, 共&N頁"
#
#    wb_out.SaveAs(school_filename)
#
#    Excel_output.Application.Quit()
