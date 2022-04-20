import hashlib
import os
import shutil
import tempfile
import time
from pathlib import Path
from typing import Iterable
from .excel_compilation import is_locked
from .py7z import pack


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
            if Path(path).is_dir():
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
