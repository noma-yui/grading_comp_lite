import datetime
import zoneinfo
import pptx


def get_creator_lastmodify(presentation):
    """Returns the creator and lastmodifiedby.

    Returns the creator and lastmodifiedby of the file.
    ファイルの作成者、最終更新者を返します。

    Args:
        document (Document): Document instance of the python-pptx

    Returns:
        (creator, lastmodifiedby) : tuple of strs
            (作成者, 最終更新者)
    """
    return (presentation.core_properties.author, presentation.core_properties.last_modified_by)


def get_createtime_modifiedtime(presentation, iana_key='Asia/Tokyo'):
    """Returns the createdtime and lastmodifiedtime.

    Returns the created datetime and the lastmodified datetime of the file.
    The default timezone info is JST.
    ファイルの作成日時、最終更新日時を返します。
    デフォルトのタイムゾーンは日本標準時間です。

    Args:
        document (Document): Document instance of the python-pptx

        iana_key : str
            IANA timezone identifier

    Returns:
        (createdtime, lastmodifiedtime) : tuple of strs
            (作成者, 最終更新者)
            The datatimes are isoformat strings.
    """
    # get datetime with "Z", (UTC)
    createdtime = presentation.core_properties.created
    modifiedtime = presentation.core_properties.modified
    # ただし、時間帯情報　timezone はNULLである　つまりシステム依存の時間に見えてしまう。
    # 日本時間に変換
    # 強引にUTCと認識させ、そこから日本時間帯に変換させる
    if createdtime:
        tmp = createdtime.replace(tzinfo=datetime.timezone.utc)
        createdtimeJST = tmp.astimezone(tz=zoneinfo.ZoneInfo(key=iana_key))
    else:
        createdtimeJST = "Empty DateTime"
    if modifiedtime:
        tmp = modifiedtime.replace(tzinfo=datetime.timezone.utc)
        modifiedtimeJST = tmp.astimezone(tz=zoneinfo.ZoneInfo(key=iana_key))
    else:
        modifiedtimeJST = "Empty DateTime"
    return (createdtimeJST, modifiedtimeJST)


