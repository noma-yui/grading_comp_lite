import os
import datetime
import zoneinfo
import win32com.client
import time






def get_creator_lastmodify(document):
    """Returns the creator and lastmodifiedby.

    Returns the creator and lastmodifiedby of the file.
    ファイルの作成者、最終更新者を返します。

    Args:
        document (Document): Document instance of the python-docx

    Returns:
        (creator, lastmodifiedby) : tuple of strs
            (作成者, 最終更新者)
    """
    return (document.core_properties.author, document.core_properties.last_modified_by)


def get_createtime_modifiedtime(document, iana_key='Asia/Tokyo'):
    """Returns the createdtime and lastmodifiedtime.

    Returns the created datetime and the lastmodified datetime of the file.
    The default timezone info is JST.
    ファイルの作成日時、最終更新日時を返します。
    デフォルトのタイムゾーンは日本標準時間です。

    Args:
        document (Document): Document instance of the python-docx

        iana_key : str
            IANA timezone identifier

    Returns:
        (createdtime, lastmodifiedtime) : tuple of strs
            (作成者, 最終更新者)
            The datatimes are isoformat strings.
    """
    # get datetime with "Z", (UTC)
    createdtime = document.core_properties.created
    modifiedtime = document.core_properties.modified
    # ただし、時間帯情報　timezone はNULLである　つまりシステム依存の時間に見えてしまう。
    # 日本時間に変換
    # 強引にUTCと認識させ、そこから日本時間帯に変換させる
    if createdtime:
        tmp = createdtime.replace(tzinfo=datetime.timezone.utc)
        createdtimeJST = tmp.astimezone(tz=zoneinfo.ZoneInfo(key=iana_key))
    else:
        createdtimeJST = "Empty Datetime"
    if modifiedtime:
        tmp = modifiedtime.replace(tzinfo=datetime.timezone.utc)
        modifiedtimeJST = tmp.astimezone(tz=zoneinfo.ZoneInfo(key=iana_key))
    else:
        modifiedtimeJST = "Empty Datetime"
    return (createdtimeJST, modifiedtimeJST)



# ファイルの内容を比較してそれを保存する。
# originaldocname, newdocname, outputfilenameはファイル名
# reference https://stackoverflow.com/questions/47212459/automating-comparison-of-word-documents-using-python
def create_word_diff(original_doc_path, students_filelist, sleeptime = 10):
    """Compare two document with Word's compare document

    Args:
        original_doc_path (Path): filename for the original file
        students_filelist (list of Path): filenames submitted by students 
        sleeptime (int, optional): Wait time until the Word will be quit. Defaults to 10.
    """
    # I learned many things from https://stackoverflow.com/questions/47212459/automating-comparison-of-word-documents-using-python
    # Thanks to the authors.

    absorig = str(original_doc_path.resolve())
    app1 = win32com.client.gencache.EnsureDispatch("Word.Application")
    orig_doc = app1.Documents.Open(absorig)

    # set Word to print view
    # https://learn.microsoft.com/ja-jp/office/vba/api/word.wdviewtype
    # WdViewType 列挙 (Word)
    # 名前	値	説明
    # wdPrintView	3	印刷レイアウト表示
    app1.ActiveDocument.ActiveWindow.View.Type = 3

    for student_doc_path in students_filelist:
        student_doc = str(student_doc_path.resolve())
        cmp_doc = student_doc + "_cmp.docx"
        doc2 = app1.Documents.Open(student_doc)
        # compare
        app1.CompareDocuments(orig_doc, doc2)
        app1.ActiveDocument.SaveAs(cmp_doc)
        
        time.sleep(sleeptime)
    
    app1.Quit()




