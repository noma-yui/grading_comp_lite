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
    filepath = pathlib.Path(emlpath)
    tmpdic = {}
    with filepath.open("rb") as email_file:
        msg = email.message_from_bytes(
            email_file.read(), policy=email.policy.default)
        for key in keys:
            tmpdic[key] = msg[key]
    return tmpdic


def get_messagebody(emlpath):
    filepath = pathlib.Path(emlpath)
    with filepath.open("rb") as email_file:
        msg = email.message_from_bytes(
            email_file.read(), policy=email.policy.default)
        part = msg.get_body(preferencelist=("plain", "html"))
        charset = str(part.get_content_charset())
        retstr = part.get_payload(
            decode=True).decode(charset, errors="strict")
    return retstr


def get_attached(emlpath, outdir):
    filepath = pathlib.Path(emlpath)
    at_filenamelist = []
    with filepath.open("rb") as email_file:
        msg = email.message_from_bytes(
            email_file.read(), policy=email.policy.default)
        for part in msg.iter_attachments():
            if part.get_filename() is not None:
                at_filename = part.get_filename()
                at_filenamelist.append(at_filename)
                # charset = str(part.get_content_charset(failobj=""))
                charset = part.get_content_charset(failobj="")
                logger.debug("attached filename={}".format(at_filename))
                logger.debug("charset={}".format(charset))
                outfilename = os.path.join(outdir, at_filename)
                with open(outfilename, "bw") as fout:
                    fout.write(part.get_payload(decode=True))
                    # if charset == "":
                    #     fout.write(part.get_payload(decode=True))
                    # else:
                    #     fout.write(part.get_payload(decode=True))
    return at_filenamelist
