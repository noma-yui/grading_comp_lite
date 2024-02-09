from pathlib import Path
import os
import msoffcrypto
from msoffcrypto.format.ooxml import OOXMLFile
from logging import getLogger, StreamHandler, DEBUG, INFO
logger = getLogger(__name__)
handler = StreamHandler()
loglevel = INFO
handler.setLevel(loglevel)
logger.setLevel(loglevel)
logger.addHandler(handler)
logger.propagate = False
# reference https://qiita.com/amedama/items/b856b2f30c2f38665701
# https://docs.python.org/ja/3/howto/logging.html


def listfiles(rootdir, ext=None):
    """List all files in some directory.

    Example:
        filelist = listfiles("./", ".py")
    Args:
        rootdir (str | pathlib.Path): directory name 
        ext (str, optional): file extention which you want to list.
            Defaults to None.

    Returns:
        list(pathlib.Path): list of files
    """
    adir = Path(rootdir).absolute()

    if ext is None:
        filelist = list(adir.glob("**/*"))
    else:
        filelist = list(adir.glob("**/*{}".format(ext)))

    return list(filter(lambda p: p.is_file(), filelist))


def encript_xlsxs(rootdir, password, flag_delete=False):
    """Encript all xlsx files in a directory.

    Args:
        rootdir (str | pathlib.Path): directory where the xlsx file exist
        password (str): password to encript
        flag_delete (bool, optional): flag to delete original files.
            Defaults to False. 

    In order to test this function
    copy files in "sampledata/exceldir1" to "sampledata/exceldir1_enc"
    and run 
    encript_xlsxs("sampledata/exceldir1_enc", password="hoge", flag_delete=True)

    """
    # list all xlsx files
    filelist = listfiles(rootdir, ext=".xlsx")

    for filepath in filelist:
        logger.debug("Encript " + str(filepath))
        # 暗号化
        try:
            f = filepath.open("rb")
            file = msoffcrypto.format.ooxml.OOXMLFile(f)

            # 保存
            outfilename = filepath.stem + "_enc.xlsx"
            outfilepath = filepath.parent / Path(outfilename)
            with outfilepath.open("wb") as f_enc:
                file.encrypt(password, f_enc)

            f.close()
            if flag_delete:
                # 元データ削除
                logger.info("Delete " + str(filepath))
                filepath.unlink()
        except:
            logger.warning("Something is wrong for " + str(filepath))


def decript_xlsxs(rootdir, password, flag_delete=False):
    """Decript all xlsx files in a directory.

    Args:
        rootdir (str | pathlib.Path): directory where the encripted files exist
        password (str): password to decript
        flag_delete (bool, optional): flag to delete original encripted files.
            Defaults to False.

    In order to test this function
    copy files in "sampledata/exceldir1_enc" to "sampledata/exceldir1_enc_dec"
    and run 
    decript_xlsxs("sampledata/exceldir1_enc_dec", password="hoge", flag_delete=True)
    """
    # list all xlsx files
    filelist = listfiles(rootdir, ext=".xlsx")

    for filepath in filelist:
        logger.debug("Decript " + str(filepath))
        # 暗号化解除
        try:
            f_enc = filepath.open("rb")
            file = msoffcrypto.OfficeFile(f_enc)
            file.load_key(password=password)

            # 保存
            outfilename = filepath.stem + "_dec.xlsx"
            outfilepath = filepath.parent / Path(outfilename)
            with outfilepath.open("wb") as f_dec:
                file.decrypt(f_dec)

            f_enc.close()
            if flag_delete:
                # 元データ削除
                logger.info("Delete " + str(filepath))
                filepath.unlink()
        except:
            logger.warning("something is wrong for " + str(filepath))
