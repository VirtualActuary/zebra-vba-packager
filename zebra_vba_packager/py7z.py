from contextlib import contextmanager
import deprecation
from pathlib import Path
import os
import shutil
from py7zr import pack_7zarchive, unpack_7zarchive
import tempfile


@contextmanager
def in_temp_directory():
    old_cwd = os.getcwd()
    with tempfile.TemporaryDirectory() as f:
        os.chdir(f)
        yield Path(f)
        os.chdir(old_cwd)


def _pack_7zarchive(path, *args, **kwargs):
    pack_7zarchive(path, *args, **kwargs)
    return str(path) + ".7z"


shutil.register_archive_format("7zip", _pack_7zarchive, description="7zip archive")
shutil.register_unpack_format("7zip", [".7z"], unpack_7zarchive)


@deprecation.deprecated(
    deprecated_in="0.0.10",
    removed_in="1.0",
    details="Use shutil.unpack_archive instead",
)
def unpack(path, targetdir, fullpaths=True):
    """
    Unpack the archive at path to targetdir. Return True if succesful.
    """
    if not fullpaths:
        raise ValueError(f"Parameter fullpaths={fullpaths} not supported anymore.")
    shutil.unpack_archive(path, targetdir)


@deprecation.deprecated(
    deprecated_in="0.0.10", removed_in="1.0", details="Use shutil.make_archive instead"
)
def pack(path, archivepath, compression_type=None, password=None):
    if password:
        raise ValueError(f"Parameter password={password} not supported anymore.")

    ext = os.path.splitext(archivepath)[-1].lower()

    if compression_type is None:
        for achive_format, _ in shutil.get_archive_formats():
            with in_temp_directory():  # don't trust dry runs
                tstout = shutil.make_archive("_", achive_format, "_", dry_run=True)
                tstext = os.path.splitext(tstout)[-1]
                if tstext == ext:
                    compression_type = achive_format
                    break

    if compression_type is None:
        raise ValueError(f"Could not determine archive format for {ext} extension")

    with in_temp_directory():  # don't trust dry runs
        outext = os.path.splitext(
            shutil.make_archive("_", compression_type, "_", dry_run=True)
        )[-1]

    if ext == outext:
        archivepath = Path(archivepath).with_suffix("")

    shutil.make_archive(archivepath, compression_type, path)
