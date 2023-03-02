import os
import shutil
from copy import deepcopy

from download import download

from .fix_repo_history_comment import fix_repo_history_comment, add_repo_history_comment
from . import util
from .fix_module_name_length_limitation import (
    _ModnamePair,
    fix_module_name_length_limitation,
)
from .bas_combining import compile_bas_sources_into_single_file
from .downloader import git_download
from .util import file_md5, first
from .vba_renaming import (
    NameTransformer,
    cls_renaming_dict,
    do_renaming,
    bas_create_namespaced_classes,
    vba_module_name,
    write_tokens,
)
import inspect
from typing import Union, List, Dict, Callable, Tuple

from dataclasses import dataclass

from pathlib import Path

import hashlib
import uuid
import tempfile

from pathvalidate import sanitize_filename

from .py7z import unpack
from .vba_tokenizer import tokenize
from .fix_casing import fix_casing


def strhash(x):
    return hashlib.md5(x.encode("utf-8")).hexdigest()


def caller_path(frame_info: inspect.FrameInfo):
    if (path := frame_info.frame.f_globals.get("__file__", None)) is not None:
        return Path(os.path.abspath(path)).resolve()


def caller_id(frame_info: inspect.FrameInfo):
    if callpath := caller_path(frame_info) is None:
        callpath = uuid.uuid4()

    return strhash(f"{callpath}@{frame_info.frame.f_lineno}")


@dataclass(init=True)
class Source:
    pre_process: Callable = None

    git_source: str = None
    git_rev: str = None
    git_add_version_comment: bool = None

    url_source: str = None
    url_md5: str = None

    path_source: Union[str, Path] = None

    glob_extract: Union[str, List[str]] = None
    glob_include: Union[str, List[str]] = "**/*"
    glob_exclude: Union[str, List[str]] = None

    mid_process: Callable = None

    combine_bas_files: Union[bool, str] = False

    auto_bas_namespace: bool = True
    auto_cls_rename: bool = True
    rename_overwrites: Union[
        Dict[str, str], List[Tuple[Union[Callable, str], Union[Callable, str]]]
    ] = None

    post_process: Callable = None

    def __post_init__(self):
        self.glob_exclude = util._str_parameter_to_list(self.glob_exclude)
        self.glob_exclude.append("**/z__NameSpaces.bas")

        if (
            sum(
                [
                    i is not None
                    for i in (self.git_source, self.url_source, self.path_source)
                ]
            )
            > 1
        ):
            raise ValueError(
                "Not more than one of git_source/url_source/path_source may be filled in"
            )

        # noinspection PyProtectedMember
        self.caller_path = caller_path(inspect.stack()[2])
        self.caller_id = caller_id(inspect.stack()[2])
        self.uid = str(uuid.uuid4())[:8]

        link = [
            i
            for i in (self.git_source, self.url_source, self.path_source, self.uid)
            if i is not None
        ][0]
        fname = sanitize_filename(
            str(link).replace("\\", "/").rstrip("/").split("/")[-1]
        )

        self.temp_downloads = Path(tempfile.gettempdir()).joinpath(
            "zebra-vba-packager",
            self.caller_id[:8]
            + "-"
            + strhash(
                str(self.git_source) + str(self.url_source) + str(self.path_source)
            )[:8]
            + "-"
            + fname,
        )

        self.temp_transformed = self.temp_downloads.parent.joinpath(
            self.temp_downloads.name + "-transformed"
        )
        if self.url_source:
            self.temp_downloads_file = Path(str(self.temp_downloads) + "-file-download")
            util.dir_touch(self.temp_downloads_file)

        util.dir_touch(self.temp_downloads)
        util.dir_touch(self.temp_transformed)


class Config:
    def __init__(self, *sources, casing=None, casing_exclusions=None):
        # noinspection PyProtectedMember
        self.caller_path = caller_path(inspect.stack()[1])
        self.caller_id = caller_id(inspect.stack()[1])
        self.sources = sources
        self.output_dir = None
        self.casing = casing
        self.casing_exclusions = casing_exclusions

    def run(self, output_dir=None):
        util.delete_old_files_in_tempdir()
        for source in self.sources:
            if source.pre_process is not None:
                source.pre_process(source)

            ltype, link = [
                (i, j)
                for (i, j) in {
                    "git": source.git_source,
                    "url": source.url_source,
                    "path": source.path_source,
                }.items()
                if j is not None
            ][0]

            # Get the files from the sources
            if ltype == "git":
                git_download(link, source.temp_downloads, source.git_rev)

            elif ltype == "url":
                # Archive sensitive unpacking
                is_archive = [
                    True
                    for i in [".zip", ".tar", ".7z", ".gz"]
                    if str(source.temp_downloads.name).lower().endswith(i)
                ]

                dlfile = source.temp_downloads_file.joinpath(source.temp_downloads.name)

                if not (dlfile.is_file() and file_md5(dlfile) == source.url_md5):
                    download(link, dlfile, replace=True)
                    if dlfile.is_file():
                        print(f"MD5 {file_md5(dlfile)} for link {link}")

                if not is_archive:
                    shutil.copy2(
                        dlfile,
                        source.temp_downloads.joinpath(source.temp_downloads.name),
                    )

                elif str(dlfile).lower().endswith(".tar.gz"):
                    unpack(dlfile, (dlgz := str(dlfile) + "-tmpunpack"))
                    for i in Path(dlgz).rglob("*"):
                        if str(i).endswith(".tar"):
                            unpack(i, source.temp_downloads)
                    util.rmtree(dlgz)

                else:
                    unpack(dlfile, source.temp_downloads)

            elif ltype == "path":
                util.rmtree(source.temp_downloads, ignore_errors=True)
                os.makedirs(source.temp_downloads, exist_ok=True)
                # Files are copied one by one because of an edge case where the destination directory
                # wasn't completely deleted (potentially when util.rmtree has ignore_errors=True)
                for i in Path(link).glob("*"):
                    ii = source.temp_downloads.joinpath(i.name)
                    os.makedirs(ii.parent, exist_ok=True)

                    if i.is_file():
                        shutil.copy(i, ii)
                    else:
                        shutil.copytree(i, ii)

            # Do the unpacking thing
            util.unpack_globs(source.glob_extract, source.temp_downloads)

            # Include/Exclude patterns
            file_matches = util.get_matching_file_patterns(
                source.temp_downloads, source.glob_include, source.glob_exclude
            )

            util.rmtree(source.temp_transformed, ignore_errors=True)
            os.makedirs(source.temp_transformed, exist_ok=True)
            for i in file_matches:
                reli = i.relative_to(source.temp_downloads)
                dst = source.temp_transformed.joinpath(reli)
                os.makedirs(dst.parent, exist_ok=True)
                shutil.copy2(i, source.temp_transformed.joinpath(reli))

            # mid process
            if source.mid_process is not None:
                source.mid_process(source)

            renames = deepcopy(source.rename_overwrites)
            if renames is None:
                renames = {}

            # Do variable renaming
            rename_transform = NameTransformer(renames)

            if source.auto_cls_rename:
                d = cls_renaming_dict(source.temp_transformed, rename_transform)

                if isinstance(renames, dict):
                    renames.update(d)
                else:
                    renames = list(renames) + [(i, j) for (i, j) in d.items()]

                rename_transform = NameTransformer(renames)

            do_renaming(source.temp_transformed, rename_transform)

            if source.combine_bas_files:
                name = (
                    source.combine_bas_files
                    if isinstance(source.combine_bas_files, str)
                    else None
                )
                sources = {}
                for f in Path(source.temp_transformed).rglob("*.bas"):
                    sources[f] = util.read_txt(f)

                if len(sources):
                    txt = compile_bas_sources_into_single_file(
                        sources, module_name=name
                    )
                    for i in sources:
                        i.unlink()

                    util.write_txt(first(sources), txt)

            if source.auto_bas_namespace:
                bas_create_namespaced_classes(source.temp_transformed)

            if source.git_add_version_comment or source.git_add_version_comment is None:
                fix_repo_history_comment(source.temp_transformed)
                if ltype == "git":
                    add_repo_history_comment(
                        source.temp_transformed, link, str(source.git_rev)
                    )

            # post process
            if source.post_process is not None:
                source.post_process(source)

        if output_dir is None and self.output_dir is None:
            self.output_dir = Path(tempfile.gettempdir()).joinpath(
                "zebra-vba-packager", self.caller_id[:8], "output"
            )
        if output_dir is None:
            output_dir = self.output_dir

        output_dir = Path(output_dir)

        util.rmtree(output_dir, ignore_errors=True)
        os.makedirs(output_dir, exist_ok=True)

        for source in self.sources:
            for i in source.temp_transformed.rglob("*"):
                if i.is_dir():
                    continue

                reli = i.relative_to(source.temp_transformed)
                if str(reli).lower()[-4:] in (".cls", ".bas"):
                    modname = vba_module_name(tokenize(util.read_txt(i)))
                    dst = output_dir.joinpath(modname + str(reli).lower()[-4:])
                else:
                    dst = output_dir.joinpath(reli)

                os.makedirs(dst.parent, exist_ok=True)
                shutil.copy2(i, dst)

        # Write namespace declarations
        namespace_declarations = [
            'Attribute VB_Name = "z__NameSpaces"',
            "' This file is generated by Zebra VBA Packager https://github.com/AutoActuary/zebra-vba-packager",
        ]
        namespace_declarations_not_empty = False

        fix_module_name_length_limitation(output_dir)

        fix_repo_history_comment(output_dir)

        if self.casing is not None or self.casing_exclusions is not None:
            fix_casing(output_dir, self.casing, self.casing_exclusions)

        for i in output_dir.rglob("*.cls"):
            if i.name.startswith("z__") and i.name.lower().endswith(".cls"):
                txt = util.read_txt(i)

                namespace_declarations_not_empty = True
                mpair = _ModnamePair.from_str(txt)
                nspacename = mpair.namespace.rstrip("__")  # rstrip legacy __ suffix
                namespace_declarations.append(
                    f"Public {nspacename} As New {mpair.input_modname}"
                )

        if namespace_declarations_not_empty:
            write_tokens(
                output_dir.joinpath("z__NameSpaces.bas"),
                tokenize("\n".join(namespace_declarations)),
            )
