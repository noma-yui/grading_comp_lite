import sys
import os
import filecmp
import shutil
import pathlib

# append import path
pardir = os.path.dirname(os.path.abspath(__file__))
utildir = os.path.join(pardir, "../")
sys.path.append(utildir)
import util

sampledata_dir = os.path.join(pardir, "../sampledata/email")

testfiles =[
    "email_data1.eml",
    "email_data2.eml",
    "email_data3_html.eml",
    "email_data4_attached.eml",
]
filelist = []
for item in testfiles:
    filelist.append(pathlib.Path(os.path.join(sampledata_dir,item)))


#######################################
print("#######################################")
print("Check email metadata. 電子メールのメタデータのチェック")
expected_meta = {
    "email_data1.eml": {
        "From": "sender_mail@example.com",
        "To": "reciepient_mail@example.com",
        "Cc": "reciepient2_mail@example.com",
        "Subject": "hello subject"
    },
    "email_data2.eml": {
        "From": "sender_mail@example.com",
        "To": "受信者１reciepient_mail@example.com",
        "Cc": "受信者２reciepient2_mail@example.com",
        "Subject": "text mail"
    },
    "email_data3_html.eml": {
        "From": "sender_mail@example.com",
        "To": "受信者１reciepient_mail@example.com",
        "Cc": "受信者２reciepient2_mail@example.com",
        "Subject": "html mail"
    },
    "email_data4_attached.eml": {
        "From": "sender_mail@example.com",
        "To": "受信者１reciepient_mail@example.com",
        "Cc": "受信者２reciepient2_mail@example.com",
        "Subject": "attached mail"
    }
}
for filepath in filelist:
    print(str(filepath.resolve()))
    retval = util.emailutil.get_header(
        filepath, ["From", "To", "Cc", "Subject"])
    expectedvals = expected_meta[filepath.name]
    assert (retval == expectedvals)
    print("OK")

#######################################
print("#######################################")
print("Check email message body. 電子メールの本文をチェックする")
expected_body = {
    "email_data1.eml": "This is a message body.",
    "email_data2.eml": "これはテキスト形式",
    "email_data3_html.eml": "This is an html mail.",
    "email_data4_attached.eml": "This is a message body.\r\nここがメッセージ本文です。\r\n添付ファイルがあります。",
}

for filepath in filelist:
    print(str(filepath.resolve()))
    retval = util.emailutil.get_messagebody(filepath)
    expectedvals = expected_body[filepath.name]
    assert (retval.strip() == expectedvals)
    print("OK")

#######################################
print("#######################################")
print("Check attached files. 電子メールの添付ファイルをチェックする")

# 添付ファイルのオリジナルファイルのディレクトリ
attach_orig_dir = os.path.join(pardir, "../sampledata/email/attached_original/")

# 添付ファイル出力用ディレクトリ
attach_out_dir = os.path.join(pardir, "../sampledata/email/attached_test/")
# directory があれば消してからディレクトリを作る
if os.path.exists(attach_out_dir):
    # 一度全部消す
    shutil.rmtree(attach_out_dir)
    os.mkdir(attach_out_dir)
else:
    os.mkdir(attach_out_dir)

for filepath in filelist:
    retval = util.emailutil.get_attached(filepath, attach_out_dir)

dirdiff = filecmp.dircmp(attach_orig_dir, attach_out_dir)
assert (not dirdiff.left_only)
assert (not dirdiff.right_only)
assert (not dirdiff.diff_files)

print("OK")
