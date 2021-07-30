import os
import shutil
from copy import deepcopy

from download import download

from .downloader import git_download
from .util import file_md5
from .vba_renaming import NameTransformer, cls_renaming_dict, do_renaming, bas_create_namespaced_classes, \
    vba_module_name, write_tokens
import inspect
from typing import Union, List, Dict, Callable, Tuple

import locate
import locate.locate
from dataclasses import dataclass

from pathlib import Path

import hashlib
import uuid
import tempfile

from pathvalidate import sanitize_filename

from .py7z import unpack
from .vba_tokenizer import tokenize


def strhash(x):
    return hashlib.md5(x.encode("utf-8")).hexdigest()

def callerpath(stackframe):
    return Path(str(locate.locate._file_path_from_stack_frame(stackframe)).replace("/", "\\")).resolve()


def str_parameter_to_list(x):
    if isinstance(x, str):
        return [x]
    if x is None:
        return []
    return x


@dataclass(init=True)
class Source:
    pre_process: Callable = None

    git_source: str = None
    url_source: str = None
    path_source: str = None
    git_rev: str = None
    url_md5: str = None

    glob_extract: Union[str, List[str]] = None

    glob_include: Union[str, List[str]] = "**/*"
    glob_exclude: Union[str, List[str]] = None

    mid_process: Callable = None

    auto_bas_namespace: bool = True
    auto_cls_rename: bool = True
    rename_overwrites: Union[Dict[str, str],
                             List[Tuple[Union[Callable, str],
                                        Union[Callable, str]]]] = None

    post_process: Callable = None

    def __post_init__(self):
        if sum([i is not None for i in (self.git_source, self.url_source, self.path_source)]) > 1:
            raise ValueError("Not more than one of git_source/url_source/path_source may be filled in")

        # noinspection PyProtectedMember
        self.caller = callerpath(inspect.stack()[2].frame)
        self.uid = str(uuid.uuid4())[:8]

        link = [i for i in (self.git_source, self.url_source, self.path_source, self.uid) if i is not None][0]
        fname = sanitize_filename(str(link).replace("\\", "/").rstrip("/").split("/")[-1])

        self.temp_downloads = Path(tempfile.gettempdir()).joinpath(
                                "zebra-vba-packager",
                                strhash(str(self.caller))[:8],
                                strhash(str(self.git_source)+str(self.url_source)+str(self.path_source))[:8]+"-"+fname)

        self.temp_transformed = self.temp_downloads.parent.joinpath(self.temp_downloads.name+"-transformed")

        os.makedirs(self.temp_downloads, exist_ok=True)
        os.makedirs(self.temp_transformed, exist_ok=True)


class Config:
    def __init__(self, *sources):
        # noinspection PyProtectedMember
        self.caller = callerpath(inspect.stack()[1].frame)
        self.sources = sources
        self.output_dir = None

    def run(self, output_dir=None):
        for source in self.sources:

            if source.pre_process is not None:
                source.pre_process(source)

            ltype, link = [(i, j) for (i, j) in {"git": source.git_source,
                                                 "url": source.url_source,
                                                 "path": source.path_source,
                                                 None: None}.items() if j is not None][0]

            # Get the files from the sources
            if ltype == "git":
                git_download(link, source.temp_downloads, source.git_rev)

            elif ltype == "url":
                # Archive sensitive unpacking
                is_archive = [str(source.temp_downloads.name).lower().endswith(i) for i in
                              [".zip", ".tar", ".7z", ".rar", ".gz"]]
                temp_downloads_file = Path(str(source.temp_downloads)+"-file-download")

                dlfile = temp_downloads_file.joinpath(source.temp_downloads.name)

                if not(dlfile.is_file() and file_md5(dlfile) == source.url_md5):
                    download(link, dlfile, replace=True)

                if not is_archive:
                    shutil.copy2(dlfile, source.temp_downloads.joinpath(source.temp_downloads.name))

                elif str(dlfile).lower().endswith(".tar.gz"):
                    unpack(dlfile, (dlgz := str(dlfile)+"-tmpunpack"))
                    for i in Path(dlgz).rglob("*"):
                        if str(i).endswith(".tar"):
                            unpack(i, source.temp_downloads)
                    shutil.rmtree(dlgz)

                elif sum([str(dlfile).lower().endswith(i) for i in [".zip", ".tar", ".7z", ".rar", ".gz"]]):
                    unpack(dlfile, source.temp_downloads)

            elif ltype == "path":
                shutil.rmtree(source.temp_downloads, ignore_errors=True)
                os.makedirs(source.temp_downloads, exist_ok=True)
                for i in Path(link).glob("*"):
                    ii = source.temp_downloads.joinpath(i.name)
                    os.makedirs(ii.parent, exist_ok=True)

                    if i.is_file():
                        shutil.copy(i, ii)
                    else:
                        shutil.copytree(i, ii)

            # Do the unpacking thing
            for glob in str_parameter_to_list(source.glob_extract):
                for i in Path(source.temp_downloads).glob(glob):
                    unpack(i, i.parent.joinpath(i.name+"-unpack"))

            # Include/Exclude patterns
            file_matches = {}
            def match_pop(key):
                if key in file_matches:
                    file_matches.pop(key)

            for glob in str_parameter_to_list(source.glob_include):
                for i in Path(source.temp_downloads).glob(glob):
                    i = i.resolve()
                    if i.is_dir():
                        for j in i.rglob("*"):
                            file_matches[j.resolve()] = None
                    else:
                        file_matches[i] = None

            for glob in str_parameter_to_list(source.glob_exclude):
                for i in Path(source.temp_downloads).glob(glob):
                    i = i.resolve()
                    if i.is_dir():
                        for j in i.rglob("*"):
                            match_pop(j.resolve())
                    else:
                        match_pop(i)

            shutil.rmtree(source.temp_transformed, ignore_errors=True)
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

            if source.auto_bas_namespace:
                bas_create_namespaced_classes(source.temp_transformed)

            # post process
            if source.post_process is not None:
                source.post_process(source)

        if output_dir is None and self.output_dir is None:
            self.output_dir = Path(tempfile.gettempdir()).joinpath("zebra-vba-packager",
                                                                   strhash(str(self.caller))[:8],
                                                                   "output")
        if output_dir is None:
            output_dir = self.output_dir

        output_dir = Path(output_dir)

        shutil.rmtree(output_dir, ignore_errors=True)
        os.makedirs(output_dir, exist_ok=True)

        for source in self.sources:
            for i in source.temp_transformed.rglob("*"):
                if i.is_dir():
                    continue

                reli = i.relative_to(source.temp_transformed)
                if str(reli).lower()[-4:] in ('.cls', '.bas'):
                    modname = vba_module_name(tokenize(i.open().read()))
                    dst = output_dir.joinpath(modname+str(reli).lower()[-4:])
                else:
                    dst = output_dir.joinpath(reli)

                os.makedirs(dst.parent, exist_ok=True)
                shutil.copy2(i, dst)

        # Write namespace declarations
        namespace_declarations = [
            'Attribute VB_Name = "z__NameSpaces"',
            "' This file is generated by Zebra VBA Packager https://github.com/AutoActuary/zebra-vba-packager"]

        for i in output_dir.rglob("*"):
            if str(i).lower()[-4:] in ('.cls'):
                if i.name.startswith("z__") and i.name.lower().endswith("__.cls"):
                    modname = i.name[:-4]
                    nspacename = modname.lstrip("z__").rstrip("__")
                    namespace_declarations.append(f"Public {nspacename} As New {modname}")

        write_tokens(
            output_dir.joinpath("z__NameSpaces.bas"),
            tokenize('\n'.join(namespace_declarations))
        )

        # Write Metadata file:
        metadata_declarations = [
            'Attribute VB_Name = "z__MetaData"',
            "'This Library is generated by Zebra VBA Packager https://github.com/AutoActuary/zebra-vba-packager",
        ]
        if self.caller is not None:
            metadata_declarations.extend(
                ["'The following file was used to produce all z-outputs:",
                 "",
                f"'****************************************************** {self.caller.name}"]
            )
            with open(self.caller) as f:
                metadata_declarations.extend(
                    ["'"+i for i in f.read().split("\n")]
                )
        write_tokens(
            output_dir.joinpath("z__MetaData.bas"),
            tokenize('\n'.join(metadata_declarations))
        )
