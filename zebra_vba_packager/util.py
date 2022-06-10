import datetime
import hashlib
import os
import shutil
import stat
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Iterable, Union, Callable, Any
from .excel_compilation import is_locked
from .py7z import pack, unpack


def file_md5(fname: Path):
    with open(fname, "rb") as f:
        file_hash = hashlib.md5()
        while chunk := f.read(8192):
            file_hash.update(chunk)
    return file_hash.hexdigest()


def first(iterable: Iterable):
    for i in iterable:
        return i

    raise StopIteration("Empty iterator")


def to_unix_line_endings(s):
    return s.replace("\r", "")


def to_dos_line_endings(s):
    return to_unix_line_endings(s).replace("\n", "\r\n")


def backup_last_50_paths(backup_dir, path, check_lock=True):
    path = Path(path)
    if check_lock and is_locked(path):
        raise RuntimeError(f"Path '{path}' is locked and cannot be overwritten.")

    if path.is_dir():
        # Backup the directory
        with tempfile.TemporaryDirectory() as outdir:
            zipname = Path(outdir).joinpath(path.name + ".zip")
            shutil.make_archive(zipname.with_suffix(""), "zip", path)
            return backup_last_50_paths(backup_dir, zipname, check_lock=check_lock)

    os.makedirs(backup_dir, exist_ok=True)

    keep = sorted(backup_dir.glob("*"))[-50:]
    for i in backup_dir.glob("*"):
        if i not in keep:
            os.remove(i)

    shutil.copy2(
        path, backup_dir.joinpath(f"{time.strftime('%Y-%m-%d--%H-%M-%S')}--{path.name}")
    )


def rmtree(
    path: Union[str, Path], ignore_errors: bool = False, onerror: Callable = None
) -> None:
    """
    Mimicks shutil.rmtree, but add support for deleting read-only files

    >>> import tempfile
    >>> with tempfile.TemporaryDirectory() as tdir:
    ...     os.makedirs(Path(tdir, "tmp"))
    ...     with Path(tdir, "tmp", "f1").open("w") as f:
    ...         _ = f.write("tmp")
    ...     os.chmod(Path(tdir, "tmp", "f1"), stat.S_IREAD|stat.S_IRGRP|stat.S_IROTH)
    ...     try:
    ...         shutil.rmtree(Path(tdir, "tmp"))
    ...     except Exception as e:
    ...         print(e) # doctest: +ELLIPSIS
    ...     rmtree(Path(tdir, "tmp"))
    [WinError 5] Access is denied: '...f1'

    """

    def _onerror(_func: Callable, _path: Union[str, Path], _exc_info) -> None:
        # Is the error an access error ?
        try:
            os.chmod(_path, stat.S_IWUSR)
            _func(_path)
        except Exception as e:
            if ignore_errors:
                pass
            elif onerror is not None:
                onerror(_func, _path, sys.exc_info())
            else:
                raise

    return shutil.rmtree(path, False, _onerror)


def delete_old_files_in_tempdir():
    """
    Keep removing directories older than 8 weeks from %temp%/zebra-vba-packager
    until all of them are removed or the directory has 10 directories left.
    """
    temp_files_location = Path(tempfile.gettempdir(), "zebra-vba-packager")
    temp_files = list(temp_files_location.glob("*"))
    temp_files.sort(key=lambda x: os.path.getmtime(x))
    number_of_files = len(temp_files)

    for file in temp_files:
        if number_of_files <= 10:
            return
        modified_date = os.path.getmtime(file)
        today = datetime.datetime.today()
        old_date = datetime.timedelta(weeks=8)
        if modified_date < (today - old_date).timestamp():
            rmtree(file)
            number_of_files -= 1


def dir_touch(directory_path):
    """
    Creates a directory if it doesn't exist yet.
    If it does exist. Changes that directory's "date modified" value to now.

    Args:
        directory_path:
            Path to the directory.
    """

    if os.path.exists(directory_path):
        new_temp_file_path = Path(directory_path, str(uuid.uuid4()))
        open(new_temp_file_path, "w").close()
        os.remove(new_temp_file_path)
    else:
        os.makedirs(directory_path)


def _str_parameter_to_list(x):
    if isinstance(x, str):
        return [x]
    if x is None:
        return []
    return x


def unpack_globs(glob_extract, path):
    for glob in _str_parameter_to_list(glob_extract):
        for i in Path(path).glob(glob):
            unpack(i, i.parent.joinpath(i.name + "-unpack"))


def get_matching_file_patterns(path, glob_include, glob_exclude=None):
    file_matches = set()
    for glob in _str_parameter_to_list(glob_include):
        for i in Path(path).glob(glob):
            i = i.resolve()
            if i.is_dir():
                for j in i.rglob("*"):
                    if j.is_file():
                        file_matches.add(j.resolve())
            else:
                file_matches.add(i)

    for glob in _str_parameter_to_list(glob_exclude):
        for i in Path(path).glob(glob):
            i = i.resolve()
            if i.is_dir():
                for j in i.rglob("*"):
                    if j.is_file():
                        file_matches.discard(j.resolve())
            else:
                file_matches.discard(i)

    return file_matches
