from typing import Union
import locate
from pathlib import Path
import os
import tempfile
import shutil
from . import util
import subprocess
import uuid

_compile_vbs = locate.this_dir().joinpath("bin", "compile.vbs")
_decompile_vbs = locate.this_dir().joinpath("bin", "decompile.vbs")
_runmacro_vbs = locate.this_dir().joinpath("bin", "runmacro.vbs")
_saveasxlsx_vbs = locate.this_dir().joinpath("bin", "saveasxlsx.vbs")


def _file_is_locked(path) -> Union[bool, Exception]:
    if not Path(path).exists():
        return False

    if Path(path).is_dir():
        for i in Path(path).rglob("*"):
            if i.is_file():
                if _is_locked_using_move(i):
                    return True
    else:
        path_moved = str(path) + ".testfilemovable000000"
        for i in range(1, 1000000):
            if os.path.exists(path_moved):
                path_moved = path_moved[:-6] + ("%06d" % i)
            else:
                break

        try:
            os.rename(path, path_moved)
        except (WindowsError, PermissionError) as e:
            return e

        os.rename(path_moved, path)
        return False


def is_locked(path) -> Union[bool, Exception]:
    if not Path(path).exists():
        return False

    if Path(path).is_dir():
        for i in Path(path).rglob("*"):
            if i.is_file():
                if _file_is_locked(i):
                    return True
        return False

    else:
        return _file_is_locked(path)


def compile_xl(src_dir, dst_file=None):
    """
    >>> indir = locate.this_dir().joinpath("../test/example-xl-with-vba")
    >>> xl = indir.parent.joinpath("temporary_output/example-xl-with-vba-and-rename.xlsb")
    >>> compile_xl(indir, xl)  #doctest: +ELLIPSIS
    WindowsPath('...example-xl-with-vba-and-rename.xlsb')
    """
    if dst_file is None:
        dst_file = str(src_dir) + ".xlsb"

    src_dir = Path(src_dir)
    dst_file = Path(dst_file)

    if os.path.splitext(dst_file)[-1].lower() != ".xlsb":
        raise ValueError("Currently we only support .xlsb outputs format.")

    if err := is_locked(dst_file):
        raise err(f"File '{dst_file}' cannot be overwritten")

    with tempfile.TemporaryDirectory() as tmpdirname:
        src_dir_tmp = Path(tmpdirname).joinpath(src_dir.name)
        shutil.copytree(src_dir, src_dir_tmp)

        # Adhere to the compile.vbs strict naming convention
        xlname = src_dir_tmp.joinpath(src_dir.name + ".xlsx")
        xlrename = src_dir_tmp.joinpath(src_dir_tmp.name + ".xlsx")
        if not xlrename.is_file():
            if not xlname.is_file():
                xlname = next(src_dir_tmp.glob("*.xlsx"))
            os.rename(xlname, xlrename)

        # Ensure all .txt, .bas and .cls files have 'lrln' line endings
        for patt in ["*.txt", "*.bas", "*.cls"]:
            for f in Path(src_dir_tmp).rglob(patt):
                txt = util.read_txt(f)
                util.write_txt(f, txt)

        subprocess.check_output(
            ["cscript", "//nologo", str(_compile_vbs), str(src_dir_tmp)]
        )

        dst_file_tmp = next(Path(tmpdirname).glob("*.xlsb"))

        os.makedirs(dst_file.parent, exist_ok=True)
        shutil.copy2(dst_file_tmp, dst_file)

    return dst_file


def decompile_xl(src_file, dst_dir=None):
    """
    >>> xl = locate.this_dir().joinpath("../test/example-xl-with-vba.xlsb")
    >>> outdir = xl.parent.joinpath("temporary_output/example-xl-output")
    >>> decompile_xl(xl, outdir) #doctest: +ELLIPSIS
    WindowsPath('...example-xl-output')

    >>> list(outdir.rglob("*"))  #doctest: +ELLIPSIS
    [...example-xl-output.xlsx...thisworkbook.txt...Module1.bas...]
    """

    if dst_dir is None:
        dst_dir = os.path.splitext(src_file)[0]

    src_file = Path(src_file)
    dst_dir = Path(dst_dir)
    dst_name = dst_dir.name

    if err := is_locked(dst_dir):
        raise err(f"File '{dst_dir}' cannot be overwritten")

    with tempfile.TemporaryDirectory() as tmpdirname:
        src_file_tmp = Path(tmpdirname).joinpath(src_file.name)
        shutil.copy2(src_file, src_file_tmp)

        subprocess.check_output(
            ["cscript", "//nologo", str(_decompile_vbs), str(src_file_tmp)]
        )

        dst_dir_tmp = [i for i in Path(tmpdirname).glob("*") if i.is_dir()][0]
        xl_tmp = next(dst_dir_tmp.rglob("*.xlsx"))
        os.rename(xl_tmp, xl_tmp.parent.joinpath(dst_name + ".xlsx"))

        util.rmtree(dst_dir, ignore_errors=True)
        os.makedirs(dst_dir, exist_ok=True)
        shutil.copytree(dst_dir_tmp, dst_dir, dirs_exist_ok=True)

    return dst_dir


def runmacro_xl(src_file, macroname=None):
    src_file = Path(src_file).resolve()
    macroarg = [macroname] if not macroname is None else []
    subprocess.check_output(
        ["cscript", "//nologo", str(_runmacro_vbs), str(src_file)] + macroarg
    )


def saveas_xlsx(src_file, dst_file):
    """
    >>> xl = locate.this_dir().joinpath("../test/example-xl-with-vba.xlsb")
    >>> outfile = xl.parent.joinpath("temporary_output/example-xl-as-xlsx.xlsx")
    >>> saveas_xlsx(xl, outfile) #doctest: +ELLIPSIS
    WindowsPath('...example-xl-as-xlsx.xlsx')

    """
    src_file, dst_file = Path(src_file).resolve(), Path(dst_file).resolve()

    if err := is_locked(dst_file):
        raise err(f"File '{dst_file}' cannot be overwritten")

    with tempfile.TemporaryDirectory() as src_d:
        with tempfile.TemporaryDirectory() as dst_d:
            src_tmp = Path(src_d).joinpath(uuid.uuid4().hex + "-" + src_file.name)
            dst_tmp = Path(src_d).joinpath(uuid.uuid4().hex + "-" + dst_file.name)

            shutil.copy2(src_file, src_tmp)
            subprocess.check_output(
                [
                    "cscript",
                    "//nologo",
                    str(_saveasxlsx_vbs),
                    str(src_tmp),
                    str(dst_tmp),
                ]
            )

            if not dst_tmp.exists():
                raise RuntimeError(f"Could not save `{src_file}` as `{dst_file}`")

            os.makedirs(dst_file.parent, exist_ok=True)
            shutil.copy2(dst_tmp, dst_file)

    return dst_file
