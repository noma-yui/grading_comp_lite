import sys
import os
import docx
import pathlib


# append import path
pardir = os.path.dirname(os.path.abspath(__file__))
utildir = os.path.join(pardir, "../")
sys.path.append(utildir)
import util.wordutil

sampledata_dir = os.path.join(pardir, "../sampledata/word")
sampledata_dir_path = pathlib.Path(sampledata_dir).resolve()

test_file = "testfile1.docx"


teacher_file = "teacher_word.docx"
student_files =[
    "stu1_word.docx",
    "stu2_word.docx",
    "stu3_word.docx",
]


test_file_path = pathlib.Path(os.path.join(sampledata_dir, test_file))
# test_file_abs = os.path.join(sampledata_dir, test_file)

# teacher_file_abs = os.path.join(sampledata_dir, teacher_file)
teacher_file_path = pathlib.Path(os.path.join(sampledata_dir, teacher_file))

student_files_path = []
for item in student_files:
    tmp_path = pathlib.Path(os.path.join(sampledata_dir, item))
    student_files_path.append(tmp_path)


#######################################
print("Check creator and last editor. Wordファイルの作成者と最終更新者の取得")
document1 = docx.Document(str(test_file_path.resolve()))
creator, modifiedby = util.wordutil.get_creator_lastmodify(document1)
assert (creator == "Creator1")
assert (modifiedby == "Editor2")
print("OK")


#######################################
print("Check createdtime and last modified time. Wordファイルの作成日時と最終更新日時")
document1 = docx.Document(str(test_file_path.resolve()))
createdtime, modifiedtime = util.wordutil.get_createtime_modifiedtime(document1, iana_key='Asia/Tokyo')
expected_createdtime = "2024-01-02T12:04:05+09:00"
expected_modifiedtime = "2024-11-12T22:14:15+09:00"
assert (createdtime.isoformat() == expected_createdtime)
assert (modifiedtime.isoformat() == expected_modifiedtime)
print("OK")


#######################################
print("Compare files. ファイルの比較")
util.wordutil.create_word_diff(teacher_file_path, student_files_path, sleeptime= 10)
print("Please check the files in {}".format(str(sampledata_dir_path)))


