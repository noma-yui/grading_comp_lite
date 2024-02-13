import email
import email.policy
import os
import pathlib
from logging import getLogger, StreamHandler, DEBUG, INFO
logger = getLogger(__name__)
handler = StreamHandler()
loglevel = INFO
# loglevel = DEBUG
handler.setLevel(loglevel)
logger.setLevel(loglevel)
logger.addHandler(handler)
logger.propagate = False
# reference https://qiita.com/amedama/items/b856b2f30c2f38665701
# https://docs.python.org/ja/3/howto/logging.html


def get_header(emlpath, keys):
    """Get email header data like 'To', 'Subject' from eml file

    Retrieves email header data such as To and From
    from saved email files with .eml extension.
    .eml の拡張子を持つ保存された電子メールファイルから、
    To や From といった電子メールのヘッダーデータを取得します。
    Args:
        emlpath (str | pathlib.Path): path for the .eml file
        keys (sequence of str): keys for the header
            example : ['To', 'Cc', 'From', 'Subject']

    Returns:
        dictionary: data with the given keys
            if no data exists for the key,
              a dictionary with None value for the key is returned
    """
    filepath = pathlib.Path(emlpath)
    tmpdic = {}
    with filepath.open("rb") as email_file:
        msg = email.message_from_bytes(
            email_file.read(), policy=email.policy.default)
        for key in keys:
            tmpdic[key] = msg[key]
    return tmpdic


def get_messagebody(emlpath):
    """Get email message body from eml file

    Retrieves email message body
    from saved email files with .eml extension.
    .eml の拡張子を持つ保存された電子メールファイルから、
    電子メールの本文を取得します。
    Args:
        emlpath (str | pathlib.Path): path for the .eml file

    Returns:
        str: message body
    """
    filepath = pathlib.Path(emlpath)
    with filepath.open("rb") as email_file:
        msg = email.message_from_bytes(
            email_file.read(), policy=email.policy.default)
        part = msg.get_body(preferencelist=("plain", "html"))
        charset = str(part.get_content_charset())
        retstr = part.get_payload(
            decode=True).decode(charset, errors="strict")
    return retstr


def get_attached(emlpath, outdir=None):
    """Save all attached files

    Retrieves all attached files
    from saved email files with .eml extension.
    If outdir is None, the files are saved in the same directory as the eml file. 
    .eml の拡張子を持つ保存された電子メールファイルから、
    全ての添付ファイルを保存します。
    outdir が None の場合、添付ファイルは eml ファイルと同じディレクトリに保存されます。
    Args:
        emlpath (str | pathlib.Path): path for the .eml file
        outdir (str | pathlib.Path, optional): directory where the files will be saved
            Defaults to None.

    Returns:
        list of str: list of saved files name
    """

    filepath = pathlib.Path(emlpath)
    if outdir is None:
        outdir = filepath.parent
    at_filenamelist = []
    with filepath.open("rb") as email_file:
        msg = email.message_from_bytes(
            email_file.read(), policy=email.policy.default)
        for part in msg.iter_attachments():
            if part.get_filename() is not None:
                at_filename = part.get_filename()
                at_filenamelist.append(at_filename)
                #
                charset = part.get_content_charset(failobj="")
                logger.debug("attached filename={}".format(at_filename))
                logger.debug("charset={}".format(charset))
                #
                outfilename = os.path.join(outdir, at_filename)
                with open(outfilename, "bw") as fout:
                    fout.write(part.get_payload(decode=True))
    return at_filenamelist
