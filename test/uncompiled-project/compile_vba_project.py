from locate import this_dir, prepend_sys_path
import tempfile

with prepend_sys_path("../.."):
    from zebra_vba_packager import Source, Config, compile_xl, runmacro_xl


a = Source(
    path_source=this_dir().joinpath("VBALib"),
    glob_include=["**/*.bas", "**/*.cls"],
    auto_bas_namespace=False,
    auto_cls_rename=False,
)

b = Source(
    path_source=this_dir().joinpath("MiscF"),
    glob_include=["*.bas"],
    auto_bas_namespace=True,
    auto_cls_rename=True,
    rename_overwrites={"MiscF": "Fn"},
)

c = Source(
    path_source=this_dir().joinpath("Templategen"),
    auto_bas_namespace=True,
    auto_cls_rename=True,
)

d = Source(
    path_source=this_dir().joinpath("Templategen"),
    auto_bas_namespace=True,
    auto_cls_rename=True,
    combine_bas_files="TempGem",
)


with tempfile.TemporaryDirectory() as td:
    Config(a, b, c).run(td)
    compile_xl(td, this_dir().joinpath("ExcelOut.xlsb"))
    runmacro_xl(this_dir().joinpath("ExcelOut.xlsb"), "addEarlyBindings")

    Config(a, b, d).run(td)
    compile_xl(td, this_dir().joinpath("ExcelOutComb.xlsb"))
    runmacro_xl(this_dir().joinpath("ExcelOutComb.xlsb"), "addEarlyBindings")
