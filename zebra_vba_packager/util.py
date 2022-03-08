import hashlib
import os
import shutil
import time
from pathlib import Path
from typing import Iterable


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


def backup_last_50_files(backup_dir, file):
    os.makedirs(backup_dir, exist_ok=True)

    keep = sorted(backup_dir.glob("*"))[-50:]
    for i in backup_dir.glob("*"):
        if i not in keep:
            os.remove(i)

    try:
        shutil.copy2(file, backup_dir.joinpath(f"{time.strftime('%Y-%m-%d--%H-%M-%S')}--{Path(file).name}"))
    except:
        pass