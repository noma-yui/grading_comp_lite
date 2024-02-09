from pathlib import Path

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


