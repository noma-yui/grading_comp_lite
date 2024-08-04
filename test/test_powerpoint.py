import sys
import os
import pptx
import pathlib

# append import path
pardir = os.path.dirname(os.path.abspath(__file__))
utildir = os.path.join(pardir, "../")
sys.path.append(utildir)
import util.powerpointutil

sampledata_dir = os.path.join(pardir, "../sampledata/powerpoint")
sampledata_dir_path = pathlib.Path(sampledata_dir).resolve()

test_file = "testfile1.pptx"

test_file_path = pathlib.Path(os.path.join(sampledata_dir, test_file))


#######################################
print("Check creator and last editor. ファイルの作成者と最終更新者の取得")
presentation1 = pptx.Presentation(str(test_file_path.resolve()))
creator, modifiedby = util.powerpointutil.get_creator_lastmodify(presentation1)
assert (creator == "Creator")
assert (modifiedby == "Editor")
print("OK")


#######################################
print("Check createdtime and last modified time. Wordファイルの作成日時と最終更新日時")
presentation1 = pptx.Presentation(str(test_file_path.resolve()))
createdtime, modifiedtime = util.powerpointutil.get_createtime_modifiedtime(presentation1, iana_key='Asia/Tokyo')
expected_createdtime = "2024-01-02T12:04:05+09:00"
expected_modifiedtime = "2024-11-12T22:14:15+09:00"
assert (createdtime.isoformat() == expected_createdtime)
assert (modifiedtime.isoformat() == expected_modifiedtime)
print("OK")



