import os
from copy import deepcopy
from textwrap import dedent
from typing import List

from .vba_tokenizer import tokenize, VBAToken
from pathlib import Path


class NameTransformer:

    def __init__(self, name_changes):
        # force lowercase matching
        if isinstance(name_changes, dict):
            name_changes = {i.lower(): j for (i, j) in name_changes.items()}
            self.name_changes = name_changes

        # force strings to matching functions
        else:
            self.name_changes = name_changes

    def match(self, x):
        if isinstance(self.name_changes, dict):
            return x.lower() in self.name_changes
        else:
            for (i, j) in self.name_changes:
                if isinstance(i, str):
                    if i.lower() == x.lower():
                        return True
                else:
                    if i(x):
                        return False
        return False

    def transform(self, x):
        if isinstance(self.name_changes, dict):
            return self.name_changes.get(x.lower(), x)
        else:
            for (i, j) in self.name_changes:
                if isinstance(i, str):
                    if i.lower() == x.lower():
                        if isinstance(j, str):
                            return j
                        else:
                            return j(x)
                else:
                    if i(x):
                        if isinstance(j, str):
                            return j
                        else:
                            return j(x)
        return x


def write_tokens(fname, tokens):
    with fname.open("wb") as f:
        f.write(("".join([i.text for i in tokens]).lstrip()).encode("utf-8"))


def vba_directory_mapping(dirname):

    files_to_tokens = {}
    for i in list(Path(dirname).rglob("*.bas"))+list(Path(dirname).rglob("*.cls")):
        with i.open("r") as f:
            files_to_tokens[i] = tokenize(f.read())

    return files_to_tokens


def vba_module_name(tokens: List[VBAToken]):
    # search for pattern [reserved=Attribute, wspace, reserved=VB_Name, ??, ??, ??, name=ECPTextStream]
    i = -1
    while (i := i+1) < len(tokens):
        t = tokens[i]
        if(t.type == "reserved" and t.text.lower() == "attribute"):
            j = i
            while (j := j+1) < len(tokens)-1 and tokens[j].type != "newline":
                pass

            subtokens = tokens[i:j]
            if sum([i.text.lower() == "vb_name" for i in subtokens]):
                return [i.text for i in subtokens if i.type == "name"][0]

    raise ValueError('Could not match `Attribute VB_Name = "the_module_name"`')


def get_all_names(tokens: List[VBAToken]):
    return [i.text for i in tokens if i.type == "name"]


def replace_all_names(tokens: List[VBAToken], name_transformer: NameTransformer):
    tokens = deepcopy(tokens)
    for i, t in enumerate(tokens):
        if t.type == "name":
            tokens[i].text = name_transformer.transform(t.text)

    return tokens


def cls_renaming_dict(dirname, user_name_transformer):
    vba_dir_map = vba_directory_mapping(dirname)

    d = {}
    for key, value in vba_dir_map.items():
        if str(key).lower().endswith(".cls"):
            modname = vba_module_name(value)
            if not user_name_transformer.match(modname):
                d[modname] = f"z{modname}"

    return d


def do_renaming(dirname, user_name_transformer):
    vba_dir_map = vba_directory_mapping(dirname)

    for key, value in vba_dir_map.items():
        write_tokens(key, replace_all_names(value, user_name_transformer))


def strip_bas_header(tokens):
    # Strip bass vb_name attribute
    i = -1
    while (i := i + 1) < len(tokens):
        if tokens[i].type == "reserved" and tokens[i].text.lower() == "vb_name":
            i0 = i
            i1 = i
            while (i0 := i0 - 1) > 0 and tokens[i0].type != "newline":
                pass
            while (i1 := i1 + 1) < len(tokens) - 1 and tokens[i1].type != "newline":
                pass

            subtokens = tokens[i0:i1]

            is_header = False
            for t in subtokens:
                if t.type == "reserved" and t.text.lower() == "attribute":
                    is_header = True

            if is_header:
                tokens = tokens[0:i0] + tokens[i1:]

    return tokens


def bas_create_namespaced_classes(dirname):
    module_header = dedent("""
    VERSION 1.0 CLASS
    BEGIN
      MultiUse = -1  'True
    END
    Attribute VB_Name = "__modulename__"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = False
    Attribute VB_PredeclaredId = False
    Attribute VB_Exposed = True
    """).lstrip()

    vba_dir_map = vba_directory_mapping(dirname)

    d = {}
    for filename, tokens in vba_dir_map.items():
        if str(filename).lower().endswith(".bas"):
            modname = vba_module_name(tokens)

            modname_new = f"z__{modname}__"
            modhead = tokenize(module_header.replace("__modulename__", modname_new))
            tokens = modhead + strip_bas_header(tokens)

            os.remove(filename)
            newpath = filename.parent.joinpath(modname_new+".cls")
            write_tokens(newpath, tokens)
