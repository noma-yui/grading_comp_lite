"""This is a sample script that uses this utility to score assignments.
これはこのツールを使って課題の採点をする方法のサンプルスクリプトです。
"""


import sys
import os
import openpyxl
import pathlib


# append import path
pardir = os.path.dirname(os.path.abspath(__file__))
utildir = os.path.join(pardir, "../")
sys.path.append(utildir)
import util.emailutil

datadir = pathlib.Path(os.path.join(pardir, "../sampledata/email/"))

studentresult = []

## student1
filename = os.path.join(datadir, "student1/assignment1.eml")
tmpdic = {}
tmpdic["studentid"] = "student1"
headerdata = util.emailutil.get_header(
        filename, ["From", "To", "Cc", "Subject"])
### check sender
if "student1@example.com" in headerdata["From"]:
    tmpdic["sender"] = True
else:
    tmpdic["sender"] = False
### check recipient
if "teacher@example.com" in headerdata["To"]:
    tmpdic["recipient"] = True
else:
    tmpdic["recipient"] = False
### check cc
if "teaching_assistant@example.com" in headerdata["Cc"]:
    tmpdic["TA"] = True
else:
    tmpdic["TA"] = False
### check body
messagedata = util.emailutil.get_messagebody(filename)
if messagedata.strip():
    tmpdic["bodyexist"] = True
else:
    tmpdic["bodyexist"] = False
studentresult.append(tmpdic)

## student2
filename = os.path.join(datadir, "student2/Assignment1.eml")
tmpdic = {}
tmpdic["studentid"] = "student2"
headerdata = util.emailutil.get_header(
        filename, ["From", "To", "Cc", "Subject"])
### check sender
if "student2@example.com" in headerdata["From"]:
    tmpdic["sender"] = True
else:
    tmpdic["sender"] = False
### check recipient
if "teacher@example.com" in headerdata["To"]:
    tmpdic["recipient"] = True
else:
    tmpdic["recipient"] = False
### check cc
if "teaching_assistant@example.com" in headerdata["Cc"]:
    tmpdic["TA"] = True
else:
    tmpdic["TA"] = False
### check body
messagedata = util.emailutil.get_messagebody(filename)
if messagedata.strip():
    tmpdic["bodyexist"] = True
else:
    tmpdic["bodyexist"] = False
studentresult.append(tmpdic)

## result
print(studentresult)