"""This is a sample script that uses this utility to score assignments.
これはこのツールを使って課題の採点をする方法のサンプルスクリプトです。
このサンプルでは、採点結果をcsvに保存します。
"""

import sys
import os
import csv
import openpyxl

# append import path
pardir = os.path.dirname(os.path.abspath(__file__))
utildir = os.path.join(pardir, "../")
sys.path.append(utildir)
import util

filename = os.path.join(pardir, "../sampledata/exceldata3_student5.xlsx")
sheetName = "Sheet1"

# 出力ファイル
outfilename = os.path.join(os.getcwd(), "sampleout.csv")

wbdata = openpyxl.load_workbook(filename=filename, data_only=True)
wbmath = openpyxl.load_workbook(filename=filename, data_only=False)

# データだけシート
sheetData = wbdata[sheetName]
# 数式のシート
sheetMath = wbmath[sheetName]

tmpdic = {}
tmpdic["studentid"] = "student5"

# Assignment 01
# Write a formula in E4 to add the values in cells C4 and D4.			
# C4セルとD4セルの値を足し算する数式をE4セルに書きなさい。			
# 	12	34
# 正解は 46 である。
# 期待される数式は「=C4+D4」である。
valuecheck = util.excelutil.is_given_value(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E4",
                                   value=46)
tmpdic["value1"] = valuecheck
# 数式を使っているかどうか
formulacheck = util.excelutil.is_formula(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E4")
tmpdic["formula1"] = formulacheck
# 採点できないこと
#  数式が「=C4+D4」である
#  次の数式も同じ意味である。
#  「=D4+C4」「= C4  +  D4」「=SUM(C4:D4)」「=SUM(C4, D4)」「= C4 * 100 / 100 + D4 」
# これらのパターンを調べることは、できなくはないだろうが困難である。

				
# Write a formula in G9 to multiply the values in cells C9 and E9.			
# C9セルとE9セルの値を掛け算する数式をGセルに書きなさい。			
# 	12	34	
# 正解は 408 である。
# 期待される数式は「=C9*D9」である。
valuecheck = util.excelutil.is_given_value(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E9",
                                   value=408)
tmpdic["value2"] = valuecheck
# 数式を使っているかどうか
formulacheck = util.excelutil.is_formula(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E9")
tmpdic["formula2"] = formulacheck
# 採点できないこと
#  数式が「=C9*D9」である
#  次の数式も同じ意味である。
#  「=D9*C9」「= C9  *  D9」「=PRODUCT(C9:D9)」「=PRODUCT(C9, D9)」「= (C9+100)*(D9-100) +100*C9 -100*D9 + 10000」
# これらのパターンを調べることは、できなくはないだろうが困難である。				
				
# Write a formula in G14 to calculate the average of the values in cells C14 and E14.			
# C14セルとE14セルの値の平均値を計算する数式をG14セルに書きなさい。			
# 	12	34
# 正解は 23 である。
# 期待される数式は「=AVERAGE(C14:D14)」である。
valuecheck = util.excelutil.is_given_value(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E14",
                                   value=23)
tmpdic["value3"] = valuecheck
# 数式を使っているかどうか
formulacheck = util.excelutil.is_formula(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E14")
tmpdic["formula3"] = formulacheck
# 採点できないこと
#  数式が「=AVERAGE(C14:D14)」である
#  次の数式も同じ意味である。
#  「=AVERAGE(C14, D14)」「=(C14+D14)/2」「=SUM(C14:D14)/COUNT(C14:D14)」
# これらのパターンを調べることは、できなくはないだろうが困難である。						
				
# Write a formula in G19 to calculate the average of the values in cells C19 and E19 using the "AVERAGE" function.			
# C19セルとE19セルの値の平均値を計算する数式をG19セルに「AVERAGE」関数を用いて書きなさい。			
# 	12	34
# 正解は 23 である。
# 期待される数式は「=AVERAGE(C14:D14)」である。
valuecheck = util.excelutil.is_given_value(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E19",
                                   value=23)
tmpdic["value4"] = valuecheck
# 数式を使っているかどうか
formulacheck = util.excelutil.is_formula(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="E19")
tmpdic["formula4"] = formulacheck
cellnum, oknum = util.excelutil.check_func_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                  range_string="E19", func_string="AVERAGE")
tmpdic["function4"] = oknum/cellnum
# 採点できないこと
#  数式が「=AVERAGE(C19:D19)」である
#  次の数式も同じ意味である。
#  「=AVERAGE(C19, D19)」
# これらのパターンを調べることは、できなくはないだろうが困難である。


# 作成者
creator, modifiedby = util.excelutil.get_creator_lastmodify(wbmath)
tmpdic["creator"] = creator
tmpdic["modifiedby"] = modifiedby


# 出力
fields = [
    "studentid",
    "value1",
    "formula1",
    "value2",
    "formula2",
    "value3",
    "formula3",
    "value4",
    "formula4",
    "function4",
    "creator",
    "modifiedby",
]

with open(outfilename, "w", newline="", encoding='utf-8') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames = fields)
    writer.writeheader()
    writer.writerow(tmpdic)
    print("outputfile = {}".format(outfilename))