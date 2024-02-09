import sys
import os
import shutil
import openpyxl

# append import path
pardir = os.path.dirname(os.path.abspath(__file__))
utildir = os.path.join(pardir, "../")
sys.path.append(utildir)
import util

#######################################
print("List")

dirname = os.path.join(pardir, "../sampledata/exceldir1")
expected_list = ["data1.xlsx", "dir11/data11.xlsx", "dir12/data12.xlsx"]
filepathlist = util.misc.listfiles(dirname)
# convert to relative posix style name
tmplist = []
for item in filepathlist:
    tmplist.append(item.relative_to(dirname).as_posix())

assert (tmplist == expected_list)

print("List OK")

#######################################
print("List xlsx")

expected_list = ["data1.xlsx", "dir11/data11.xlsx", "dir12/data12.xlsx"]
filepathlist = util.misc.listfiles(dirname, ext=".xlsx")
# convert to relative posix style name
tmplist = []
for item in filepathlist:
    tmplist.append(item.relative_to(dirname).as_posix())

assert (tmplist == expected_list)
print("List xlsx OK")

#######################################
print("list xls")

expected_list = []
filepathlist = util.misc.listfiles(dirname, ext=".xls")
# convert to relative posix style name
tmplist = []
for item in filepathlist:
    tmplist.append(item.relative_to(dirname).as_posix())

assert (tmplist == expected_list)
print("list xls OK")

#######################################
print("Encript xlsx")

srcdirname = os.path.join(pardir, "../sampledata/exceldir1")
dstdirname = os.path.join(pardir, "../sampledata/exceldir1_enc")

abssrcdirname = os.path.abspath(srcdirname)
absdstdirname = os.path.abspath(dstdirname)

# テスト用のデータを別のフォルダーへコピーする
if os.path.exists(absdstdirname):
    # 一度全部消す
    shutil.rmtree(absdstdirname)
    os.mkdir(absdstdirname)
else:
    os.mkdir(absdstdirname)
shutil.copytree(src=abssrcdirname, dst=absdstdirname, dirs_exist_ok=True)
util.misc.encript_xlsxs(
    rootdir=absdstdirname, password="abc", flag_delete=True)

print("Please check by hand that files are encripted {}".format(absdstdirname))

#######################################
print("Decript xlsx")
srcdirname = os.path.join(pardir, "../sampledata/exceldir1_enc")
dstdirname = os.path.join(pardir, "../sampledata/exceldir1_enc_dec")

abssrcdirname = os.path.abspath(srcdirname)
absdstdirname = os.path.abspath(dstdirname)

# テスト用のデータを別のフォルダーへコピーする
if os.path.exists(absdstdirname):
    # 一度全部消す
    shutil.rmtree(absdstdirname)
    os.mkdir(absdstdirname)
else:
    os.mkdir(absdstdirname)
shutil.copytree(src=abssrcdirname, dst=absdstdirname, dirs_exist_ok=True)
util.misc.decript_xlsxs(
    rootdir=absdstdirname, password="abc", flag_delete=True)

print("Please check by hand that files are deccripted {}".format(absdstdirname))