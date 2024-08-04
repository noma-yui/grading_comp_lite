import sys
import os
import io
import openpyxl

# append import path
pardir = os.path.dirname(os.path.abspath(__file__))
utildir = os.path.join(pardir, "../")
sys.path.append(utildir)
import util.excelutil
import util.excelutil_exp


filename = os.path.join(pardir, "../sampledata/exceldata1.xlsx")
sheetName = "Sheet1"

wbdata = openpyxl.load_workbook(filename=filename, data_only=True)
wbmath = openpyxl.load_workbook(filename=filename, data_only=False)

# データだけシート
sheetData = wbdata[sheetName]
# 数式のシート
sheetMath = wbmath[sheetName]

#######################################
print("Check data print. データの出力チェック")
# util.excelutil.print_values_in_range(sheetdata=sheetData, sheetmath=sheetMath,
#                                  range_string="C12:E13")
# util.excelutil.print_formulas_in_range(sheetdata=sheetData, sheetmath=sheetMath,
#                                    range_string="C12:E13")
outstring = """11
111
122
22
222
244
"""
tmpstringio = io.StringIO()
util.excelutil.print_values_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                range_string="C12:E13",
                                out = tmpstringio)
tmpstr = tmpstringio.getvalue()

assert (tmpstr == outstring)
outstring = """11
111
=C12+D12
22
222
=C13+D13
"""
tmpstringio = io.StringIO()
util.excelutil.print_formulas_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                range_string="C12:E13",
                                out = tmpstringio)
tmpstr = tmpstringio.getvalue()
assert (tmpstr == outstring)
print("OK")


#######################################
print("Check creator and last editor. Excelファイルの作成者と最終更新者の取得")
creator, modifiedby = util.excelutil.get_creator_lastmodify(wbmath)
assert (creator == "Creator01")
assert (modifiedby == "Editor02")
print("OK")


#######################################
print("Check createdtime and last modified time. Excelファイルの作成日時と最終更新日時")
createdtime, modifiedtime = util.excelutil.get_createtime_modifiedtime(wbmath, iana_key='Asia/Tokyo')
expected_createdtime = "2024-01-02T12:04:05+09:00"
expected_modifiedtime = "2024-11-12T22:14:15+09:00"
assert (createdtime.isoformat() == expected_createdtime)
assert (modifiedtime.isoformat() == expected_modifiedtime)
print("OK")


#######################################
print("Check that a cell has a given value. 指定したセルの値がｘｘである。")
ret1 = util.excelutil.is_given_value(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="B2",
                                   value=123)
assert (ret1)
ret1 = util.excelutil.is_given_value(sheetdata=sheetData, sheetmath=sheetMath,
                                   addr="B2",
                                   value=199)
assert (not ret1)
print("OK")


#######################################
print("Check that cell data is a formula. 指定したセルが数式である")
ret1 = util.excelutil.is_formula(sheetdata=sheetData, sheetmath=sheetMath,
                                     addr="B5")
assert (ret1)
ret1 = util.excelutil.is_formula(sheetdata=sheetData, sheetmath=sheetMath,
                                     addr="A4")
assert (not ret1)
ret1 = util.excelutil.is_formula(sheetdata=sheetData, sheetmath=sheetMath,
                                     addr="B6")
assert (ret1)
print("OK")


#######################################
print("Check values in a range. 指定したセル範囲の値が指定した値かどうか")
refdata = [
    ["aaa", 1290],
    ["bbb", 456],
    ["ccc", 789],
]
(countCells, countTrue) = util.excelutil.check_values_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                           range_string="B8:C10",
                                                           values=refdata)
assert (countCells == 6)
assert (countTrue == 6)
refdata2 = [
    ["aa", 1290],
    ["bbb", 45666],
    ["c", 789],
]
(countCells, countTrue) = util.excelutil.check_values_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                           range_string="B8:C10",
                                                           values=refdata2)
assert (countCells == 6)
assert (countTrue == 3)
print("OK")


#######################################
print("Check float values in a range. 指定したセル範囲の浮動小数点数が指定した値かどうか")
refdata = [
    [123.45678],
]
diffval = 0.001
(countCells, countTrue) = util.excelutil.check_values_in_range_float(sheetdata=sheetData, sheetmath=sheetMath,
                                                           range_string="D32",
                                                           values=refdata, diffval=diffval)
assert (countCells == 1)
assert (countTrue == 1)
print("OK")


#######################################
print("Check formulas in a range. 指定したセル範囲が数式である")
(countCells, countTrue) = util.excelutil.check_num_formulas_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                                 range_string="B12:E14")
assert (countCells == 12)
assert (countTrue == 3)
(countCells, countTrue) = util.excelutil.check_num_formulas_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                                 range_string="E12:E14")
assert (countCells == 3)
assert (countTrue == 3)
print("OK")


#######################################
print("Check that formula contains a certain function name. 指定したセル範囲の数式にｘｘ関数の文字列がある")
(countCells, countTrue) = util.excelutil.check_func_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                         range_string="D21", func_string="SUM")
assert (countCells == 1)
assert (countTrue == 1)
(countCells, countTrue) = util.excelutil.check_func_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                         range_string="D21:D28", func_string="SUM")
assert (countCells == 8)
assert (countTrue == 2)
print("OK")


#######################################
print("Check that a value is an integer. 指定したセルの値が整数")
ret1 = util.excelutil.is_integer(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="D31")
assert (ret1)
ret1 = util.excelutil.is_integer(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="D32")
assert (not ret1)
ret1 = util.excelutil.is_integer(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="D33")
assert (not ret1)
print("OK")


#######################################
print("Check composite or absolute cell reference. 指定したセル範囲の数式に複合参照/絶対参照を用いている")
(countCells, countTrue) = util.excelutil.check_comp_abs_ref_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                                 range_string="C54")
assert (countCells == 1)
assert (countTrue == 0)
(countCells, countTrue) = util.excelutil.check_comp_abs_ref_in_range(sheetdata=sheetData, sheetmath=sheetMath,
                                                                 range_string="C54:C56")
assert (countCells == 3)
assert (countTrue == 2)
print("OK")


#######################################
print("Check horizontal alignment. 指定したセルの左右の配置が指定したもの")
# デフォルト
ret1 = util.excelutil_exp.is_aligned_h(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C38",
                             horizontal=None)
assert (ret1)
# 右揃え
ret1 = util.excelutil_exp.is_aligned_h(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C41",
                             horizontal="right")
assert (ret1)
# 左揃え
ret1 = util.excelutil_exp.is_aligned_h(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C43",
                             horizontal="left")
assert (ret1)
# 中央揃え
ret1 = util.excelutil_exp.is_aligned_h(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C45",
                             horizontal="center")
assert (ret1)
# 両端揃え
ret1 = util.excelutil_exp.is_aligned_h(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C47",
                             horizontal="justify")
assert (ret1)
# NG例 データは右揃え
ret1 = util.excelutil_exp.is_aligned_h(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C41",
                             horizontal="left")
assert (not ret1)
print("OK")


#######################################
print("Check vertical alignment. 指定したセルの上下の配置が指定したもの")
# デフォルト
ret1 = util.excelutil_exp.is_aligned_v(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C38",
                             vertical=None)
assert (ret1)

#中央揃え
ret1 = util.excelutil_exp.is_aligned_v(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C41",
                             vertical="center")
assert (ret1)
#上揃え
ret1 = util.excelutil_exp.is_aligned_v(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C43",
                             vertical="top")
assert (ret1)
#下揃え（デフォルト）
ret1 = util.excelutil_exp.is_aligned_v(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C45",
                             vertical=None)
assert (ret1)
#両端揃え
ret1 = util.excelutil_exp.is_aligned_v(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C47",
                             vertical="justify")
assert (ret1)
# NG例 データは中央揃え
ret1 = util.excelutil_exp.is_aligned_v(sheetdata=sheetData, sheetmath=sheetMath,
                             addr="C41",
                             vertical=None)
assert (not ret1)
print("OK")


#######################################
print("Check solid fill. 指定したセルが塗りつぶし")
ret1 = util.excelutil_exp.is_solidfill(sheetdata=sheetData, sheetmath=sheetMath,
                               addr="C37")
assert (ret1)
ret1 = util.excelutil_exp.is_solidfill(sheetdata=sheetData, sheetmath=sheetMath,
                               addr="C38")
assert (not ret1)
print("OK")


#######################################
print("Check number format. 指定したセルの「表示形式」がｘｘである")
# General 標準
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E60",
                                  number_format="General")
assert (ret1)
# Number (built-in) 数値（組み込み）
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E61",
                                  number_format="0_);[Red]\\(0\\)")
assert (ret1)
# NG
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E61",
                                  number_format="General")
assert (not ret1)
# Percentage (built-in) パーセンテージ（組み込み）
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E62",
                                  number_format="0%")
assert (ret1)
# NG
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E62",
                                  number_format="General")
assert (not ret1)
#Text (built-in) 文字列（組み込み）
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E63",
                                  number_format="@")
assert (ret1)
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E63",
                                  number_format="General")
assert (not ret1)
# Number (0 digits after decimal point) 数値（小数点以下０桁）
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E64",
                                  number_format="0_ ")
assert (ret1)
#NG
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E64",
                                  number_format="General")
assert (not ret1)
# Number (1 digits after decimal point) 数値（小数点以下１桁）
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E65",
                                  number_format="0.0_ ")
assert (ret1)
#NG
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E65",
                                  number_format="0.00_ ")
assert (not ret1)
#Number (2 digits after decimal point) 数値（小数点以下２桁）
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E66",
                                  number_format="0.00_ ")
assert (ret1)
#NG
ret1 = util.excelutil_exp.is_numberformat(sheetdata=sheetData, sheetmath=sheetMath,
                                  addr="E66",
                                  number_format="0.000_ ")
assert (not ret1)

print("OK")



