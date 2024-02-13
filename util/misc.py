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

    Files are listed recursively.
    ディレクトリ以下のファイルを再帰的に探します。
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


def makedirs(rootdir, id_sequence):
    """Make directory with the given names

    Create multiple directories under rootdir.
    The names of the directories to be created are given by id_sequence.
    It only creates the directories in id_sequence and does not delete the existing directories.
    rootdir 以下に複数のディレクトリを作成します。
    作成するディレクトリの名前は id_sequence で与えます。
    id_sequence にあるディレクトリを作成するだけで既存のディレクトリを消すことはしません。
    Args:
        rootdir (str | pathlib.Path): directory under which the directories are created
        id_sequence (list of str | sequence of str): directory names

    """
    adir = Path(rootdir).absolute().resolve()

    for id1 in id_sequence:
        dir1 = adir / id1
        os.makedirs(dir1, exist_ok=True)


def encript_xlsxs(rootdir, password, flag_delete=False):
    """Encript all xlsx files in a directory.

    Encript the encrypted xlsx file located under rootdir.
    If flag_delete is set to True, the unencrypted files will be deleted.
    rootdir 以下にある暗号化されたxlsxファイルを暗号化します。
    flag_delete を True にすると、暗号化前のファイルは削除されます。
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
            logger.warning("Something is wrong for " + str(filepath.resolve()))


def decript_xlsxs(rootdir, password, flag_delete=False):
    """Decript all xlsx files in a directory.

    Decrypt the encrypted xlsx file located under rootdir.
    If flag_delete is set to True, the encrypted files will be deleted.
    An exception will be thrown if there is an unencrypted xlsx file under rootdir.
    rootdir 以下にある暗号化されたxlsxファイルを復号します。
    flag_delete を True にすると、暗号化前のファイルは削除されます。
    rootdir 以下に暗号化されていないxlsxファイルがあると例外を出します。
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
            logger.warning("something is wrong for " + str(filepath.resolve()))
