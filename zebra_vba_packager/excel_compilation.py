import locate
from pathlib import Path
import os
import tempfile
import shutil
import subprocess

_compile_vbs = locate.this_dir().joinpath("bin", "compile.vbs")
_decompile_vbs = locate.this_dir().joinpath("bin", "decompile.vbs")
_runmacro_vbs = locate.this_dir().joinpath("bin", "runmacro.vbs")


def is_locked(path):
    if not Path(path).exists():
        return False

    path_moved = str(path) + ".testfilemovable000000"
    for i in range(1, 1000000):
        if os.path.exists(path_moved):
            path_moved = path_moved[:-6] + "%06d"%i
        else:
            break

    # test if movable
    try:
        os.rename(path, path_moved)
    except (WindowsError, PermissionError) as e:
        return e

    os.rename(path_moved, path)
    return False


def compile_xl(src_dir, dst_file=None):
    """
    >>> indir = locate.this_dir().joinpath("../test/example-xl-with-vba")
    >>> xl = indir.parent.joinpath("temporary_output/example-xl-with-vba-and-rename.xlsb")
    >>> compile_xl(indir, xl)  #doctest: +ELLIPSIS
    WindowsPath('...example-xl-with-vba-and-rename.xlsb')
    """
    if dst_file is None:
        dst_file = str(src_dir)+".xlsb"

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
        xlname = src_dir_tmp.joinpath(src_dir.name+".xlsx")
        xlrename = src_dir_tmp.joinpath(src_dir_tmp.name+".xlsx")
        if not xlrename.is_file():
            if not xlname.is_file():
                xlname = next(src_dir_tmp.glob("*.xlsx"))
            os.rename(xlname, xlrename)

        subprocess.check_output(["cscript", "//nologo", str(_compile_vbs), str(src_dir_tmp)])

        dst_file_tmp = next(Path(tmpdirname).glob("*.xlsb"))

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

        subprocess.check_output(["cscript", "//nologo", str(_decompile_vbs), str(src_file_tmp)])

        dst_dir_tmp = [i for i in Path(tmpdirname).glob("*") if i.is_dir()][0]
        xl_tmp = next(dst_dir_tmp.rglob("*.xlsx"))
        os.rename(xl_tmp, xl_tmp.parent.joinpath(dst_name+".xlsx"))

        shutil.rmtree(dst_dir, ignore_errors=True)
        os.makedirs(dst_dir, exist_ok=True)
        shutil.copytree(dst_dir_tmp, dst_dir, dirs_exist_ok=True)

    return dst_dir


def runmacro_xl(src_file, macroname=None):
    macroarg = [macroname] if not macroname is None else []
    subprocess.check_output(["cscript", "//nologo", str(_runmacro_vbs), str(src_file)]+macroarg)
