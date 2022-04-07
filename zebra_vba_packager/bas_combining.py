from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from pprint import pprint
from textwrap import indent
from typing import Dict, Union, List
from functools import reduce
import operator

from .match_tokens import match_tokens
from .vba_renaming import vba_module_name
from .util import first, to_unix_line_endings, to_dos_line_endings
from .vba_tokenizer import VBAToken, tokenize, tokens_to_str


@dataclass
class VBASectionClassifier:
    tokens: Union[List[VBAToken], None] = None
    type: Union[str, None] = None
    origin: Union[str, None] = None
    name: Union[str, None] = None
    private: Union[bool, None] = None

    def __post_init__(self):
        if self.type is None:
            self.type = "unknown"


def get_private_names(tokens: List[VBAToken]):
    """
    Examples:
        >>> tokens = tokenize('''Attribute VB_Name = "MiscArray"
        ...
        ... Option Explicit
        ...
        ... Private Declare PtrSafe Function funcA Lib "A.dll" (ByVal x As Long) as Long
        ... Private Declare Function funcB Lib "A.dll" (ByVal x As Long) as Long
        ...
        ... private const theConst = 5
        ...
        ... ' Comment
        ... Private Function Bla(arr As Variant)
        ...     Bla = True
        ... End Function
        ... Private Function Bla2(arr As Variant)
        ...     Bla2 = True
        ... End Function
        ... ''')

        >>> get_private_names(tokens)
        ['funcA', 'funcB', 'theConst', 'Bla', 'Bla2']
    """

    return [
        tokens[j-1].text for (i, j) in match_tokens(
            tokens,
            "private [static] [declare] [ptrsafe] function|sub|parameter|enum|const .*",
            on_line_start=True
    )]


def compile_code_into_sections(
        input: Union[List[VBAToken], str],
        origin: Union[str, None] = None
) -> List[VBASectionClassifier]:
    r"""
    Split VBA code into a few high-level catagories, such as `#if`, `function`, `option`, `declare`, and
    `unknown`. Extend `unknown` into more catagories as they are needed by other parts of the codebase.
    """
    tokens = input
    if not isinstance(tokens, list):
        tokens = tokenize(input)

    classifiers = [
        VBASectionClassifier(tokens, origin=origin)
    ]

    def extend_classification_list(classifiers, idx, start_marker, end_marker, params1=None, params2=None, params3=None):
        tokens = classifiers[idx].tokens

        c1 = classifiers[idx]
        c1.tokens = tokens[:start_marker]

        c2 = VBASectionClassifier(
            tokens[start_marker:end_marker],
            origin=c1.origin
        )
        c3 = VBASectionClassifier(
            tokens[end_marker:],
            origin=c1.origin
        )
        for c, p in ((c1, params1), (c2, params2), (c3, params3)):
            if p is not None:
                for attr in dir(p):
                    if not attr.startswith("_"):
                        if getattr(p, attr) is not None:
                            setattr(c, attr, getattr(p, attr))

        classifiers.insert(
            idx + 1,
            c2
        )
        classifiers.insert(
            idx + 2,
            c3
        )

    def find_section(tokens, start_match_string, end_match_string):
        i, j = next(match_tokens(tokens, start_match_string, on_line_start=True))
        k, l = [_ + j for _ in next(match_tokens(tokens[j:], end_match_string, on_line_start=True))]
        return i, l

    # If seperation
    idx = -1
    while (idx := idx + 1) < len(classifiers):
        if classifiers[idx].type != "unknown":
            continue

        tokens = classifiers[idx].tokens
        try:
            i, j = find_section(tokens, "#if", "#end\s*if")
            extend_classification_list(classifiers, idx, i, j, params2=VBASectionClassifier(type="#if"))
            idx = idx - 1
        except StopIteration:
            break

    # Function/enum seperation
    idx = -1
    while (idx := idx + 1) < len(classifiers):
        if classifiers[idx].type != "unknown":
            continue
        tokens = classifiers[idx].tokens
        try:
            pre = "[private|public] property|sub|function|enum"
            i, j = find_section(tokens, pre, r"end property|sub|function|enum")
            type_ = tokens[j-1].text.lower()
            type_ = "function" if type_ in ("sub", "property") else type_
            extend_classification_list(classifiers, idx, i, j, params2=VBASectionClassifier(type=type_))
        except StopIteration:
            break

    # Declare/option/attribute separation (ends on newline or end of file)
    idx = -1
    while (idx := idx + 1) < len(classifiers):
        if classifiers[idx].type != "unknown":
            continue
        tokens = classifiers[idx].tokens
        try:
            i, i0 = next(match_tokens(tokens, "[private|public] declare|option|attribute", on_line_start=True))
            try:
                _, j = [x + i0 for x in next(match_tokens(tokens[i0:], r"\r\n"))]
            except StopIteration:
                j = len(tokens)

            type_ = tokens[i0-1].text.lower()
            extend_classification_list(classifiers, idx, i, j, params2=VBASectionClassifier(
                type=type_,
                private=(tokens[i].text.lower() == "private")
            ))
        except StopIteration:
            break

    # Remove empty classifications
    idx = -1
    while(idx := idx+1) < len(classifiers):
        if not len(classifiers[idx].tokens):
            classifiers.pop(idx)
            idx -= 1

    return classifiers


def compile_bas_sources_into_single_file(
        sources: Dict[Union[str, Path], Union[str, Path, List[VBAToken]]],
        module_name: Union[str, None] = None
) -> str:

    sources = {key: deepcopy(val) if isinstance(val, list) else tokenize(val) for key, val in sources.items()}
    names = {key: vba_module_name(tokens) for key, tokens in sources.items()}

    # Replace private functions and consts etc. with guarded named versions of themselves
    for key, tokens in sources.items():
        privates = {i.lower(): f"{names[key]}_{i}" for i in get_private_names(tokens)}

        for t in tokens:
            if t.type == "name" and t.text.lower() in privates:
                t.text = privates[t.text.lower()]

    classifiers_sources = {}
    for key, val in sources.items():
        classifiers_sources[key] = []
        for c in compile_code_into_sections(val, origin=key):
            classifiers_sources[key].append(c)

    # Get the first entry in #if statement and rather use that as an proxy classification for the whole block
    for classifiers in classifiers_sources.values():
        for c in classifiers:
            if c.type == '#if':
                for i in compile_code_into_sections(c.tokens[3:]):
                    if i.type != 'unknown':
                        c.type = i.type
                        break

    # Sanity check to see if all sources have the same `declare` statements:
    option_statements = {i: [] for i in sources}

    for classifiers in classifiers_sources.values():
        for c in classifiers:
            if c.type == 'option':
                normcode = " ".join([i.text.lower().strip() for i in c.tokens if not i.type in ('space', 'comment')])
                option_statements[c.origin].append(normcode)

    for key in option_statements:
        option_statements[key] = sorted(option_statements[key])

    key_a = None
    for i, key in enumerate(option_statements):
        if i == 0:
            key_a = key
        else:
            if option_statements[key_a] != option_statements[key]:
                ln = '\n'
                lhs = indent(f'{key_a}:\n{indent(ln.join(option_statements[key_a]), "  ")}', "  ")
                rhs = indent(f'{key}:\n{indent(ln.join(option_statements[key]), "  ")}', "  ")
                raise ValueError(f"Options must be equal across aggregated bas files, got conflict:\n{lhs}\n{rhs}")

    # Merge classifiers with `unknown` at their top
    for classifiers in classifiers_sources.values():
        idx = -1
        while(idx := idx+1) < len(classifiers):
            if classifiers[idx].type == 'unknown':
                for j in range(idx + 1, len(classifiers)):

                    classifiers[idx+1].tokens = classifiers[idx].tokens + classifiers[idx+1].tokens
                    classifiers.pop(idx)
                    if classifiers[idx].type != 'unknown':
                        break

    classifiers = reduce(operator.concat, classifiers_sources.values())

    buckets = {
        'attribute': [],
        'option': [],
        'declare': [],
        'other': [],
        'function': []
    }

    for c in classifiers:
        # Only use one of the file's declare statements
        if c.type == 'option':
            if c.origin == first(classifiers_sources):
                buckets[c.type].append(c)
        elif c.type in buckets:
            buckets[c.type].append(c)
        else:
            buckets['other'].append(c)

    code = []
    origin = None
    for key, vals in buckets.items():
        if key == 'attribute':
            continue

        for i in vals:

            block_str = to_unix_line_endings(tokens_to_str(i.tokens)).strip()
            if block_str != "":
                block_str = block_str + "\n\n"
                if i.type != 'option':
                    if i.origin != origin:
                        code.append(f"'*************** {names[i.origin]}\n")
                        origin = i.origin

                code.append(block_str)

    if module_name is None:
        module_name = first(names.values())

    code.insert(0, f'Attribute VB_Name = "{module_name}"\n')

    return to_dos_line_endings("".join(code))

