from copy import deepcopy
from dataclasses import dataclass
from itertools import chain
from pathlib import Path
from contextlib import suppress
from textwrap import indent
from types import SimpleNamespace as SN
from typing import Dict, Union, List
from functools import reduce
import operator
from sortedcontainers import SortedDict

from .match_tokens import match_tokens
from .vba_renaming import vba_module_name
from .util import first, to_unix_line_endings
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


def get_private_renames(tokens: List[VBAToken]):
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

        >>> get_private_renames(tokens)
        ['theConst', 'Bla', 'Bla2']
    """

    return [
        tokens[j - 1].text
        for (i, j) in match_tokens(
            tokens,
            "private [static] function|sub|parameter|enum|const .*",
            on_line_start=True,
        )
    ]


def find_section(tokens, start_match_string, end_match_string):
    i, j = next(match_tokens(tokens, start_match_string, on_line_start=True))
    k, l = [
        _ + j
        for _ in next(match_tokens(tokens[j:], end_match_string, on_line_start=True))
    ]
    return i, l


def find_all_hashif_sections(tokens):
    in_ = 0
    sections = []
    for i in range(len(tokens)):
        if tokens[i].type == "#if":
            if tokens[i].text.lower() == "#if":
                if in_ == 0:
                    sections.append([i, None])
                in_ += 1
            elif "#end" in tokens[i].text.lower():
                if in_ == 1:
                    sections[-1][-1] = i + 1
                in_ -= 1

    if sections and None in sections[-1]:
        sections = sections[:-1]

    return {tuple(i): "#if" for i in sections}


def find_all_function_sections(tokens):
    sections = {}
    start = 0
    while True:
        try:
            pre = "[private|public] property|sub|function|enum"
            i, j = find_section(tokens[start:], pre, r"end property|sub|function|enum")
            sections[(start + i, start + j)] = (
                "function"
                if (type_ := tokens[start + j - 1].text.lower()) in ("sub", "property")
                else type_
            )
            start += j
        except StopIteration:
            break

    return sections


def find_all_declaration_sections(tokens):
    sections = {}
    for i, i0 in match_tokens(
        tokens,
        "[private|public] declare|option|attribute",
        on_line_start=True,
    ):
        try:
            j = i0 + next(match_tokens(tokens[i0:], r"\n"))[1]
        except StopIteration:
            j = len(tokens)

        sections[(i, j)] = tokens[i0 - 1].text.lower()

    return sections


def find_all_global_var_sections(tokens):
    sections = {
        tuple(idx): "global"
        for idx in match_tokens(
            tokens,
            "[private|public|dim] .* as [new] .*",
            on_line_start=True,
            on_line_end=True,
        )
    }

    return sections


def compile_code_into_sections(
    input: Union[List[VBAToken], str], origin: Union[str, None] = None
) -> List[VBASectionClassifier]:
    r"""
    Split VBA code into a few high-level catagories, such as `#if`, `function`, `option`, `declare`, and
    `unknown`. Extend `unknown` into more catagories as they are needed by other parts of the codebase.
    """
    tokens = input
    if not isinstance(tokens, list):
        tokens = tokenize(input)

    mixed = SortedDict()
    for f in [
        find_all_hashif_sections,
        find_all_function_sections,
        find_all_declaration_sections,
        find_all_global_var_sections,
    ]:
        mixed.update(f(tokens))

    keys = list(mixed.keys())
    keys_sets = [set(range(key[0], key[1])) for key in keys]

    # merge overlapping sections
    i = 0
    while i < len(keys) - 1:
        if keys_sets[i].intersection(keys_sets[i + 1]):
            key_new = (keys[i][0], max(keys[i][1], keys[i + 1][1]))

            mixed[key_new] = mixed.pop(keys[i])
            mixed.pop(keys[i + 1])

            keys_sets[i] = set(range(*key_new))
            keys_sets.pop(i + 1)

            keys[i] = key_new
            keys.pop(i + 1)

            continue
        i += 1

    # Fill empty gaps
    for (_, i), (j, _) in zip(
        (l := list(mixed.keys())), chain(l[1:], [(len(tokens), None)])
    ):
        if i != j:
            mixed[i, j] = "unknown"

    return [
        VBASectionClassifier(
            tokens=(t := tokens[idxes[0] : idxes[1]]),
            type=type_,
            origin=origin,
            private=type != "unknown"
            and "private" in (_.text.lower().strip() for _ in t[:2]),
        )
        for idxes, type_ in mixed.items()
    ]


def compile_bas_sources_into_single_file(
    sources: Dict[Union[str, Path], Union[str, Path, List[VBAToken]]],
    module_name: Union[str, None] = None,
) -> str:

    sources = {
        key: deepcopy(val) if isinstance(val, list) else tokenize(val)
        for key, val in sources.items()
    }
    names = {key: vba_module_name(tokens) for key, tokens in sources.items()}

    # Replace private functions and consts etc. with guarded named versions of themselves
    for key, tokens in sources.items():
        privates = {i.lower(): f"{names[key]}_{i}" for i in get_private_renames(tokens)}

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
            if c.type == "#if":
                for i in compile_code_into_sections(c.tokens[3:]):
                    if i.type != "unknown":
                        c.type = i.type
                        break

    # Sanity check to see if all sources have the same `declare` statements:
    option_statements = {i: [] for i in sources}

    for classifiers in classifiers_sources.values():
        for c in classifiers:
            if c.type == "option":
                normcode = " ".join(
                    [
                        i.text.lower().strip()
                        for i in c.tokens
                        if not i.type in ("space", "comment")
                    ]
                )
                option_statements[c.origin].append(normcode)

    for key in option_statements:
        option_statements[key] = sorted(option_statements[key])

    key_a = None
    for i, key in enumerate(option_statements):
        if i == 0:
            key_a = key
        else:
            if option_statements[key_a] != option_statements[key]:
                ln = "\n"
                lhs = indent(
                    f'{key_a}:\n{indent(ln.join(option_statements[key_a]), "  ")}', "  "
                )
                rhs = indent(
                    f'{key}:\n{indent(ln.join(option_statements[key]), "  ")}', "  "
                )
                raise ValueError(
                    f"Options must be equal across aggregated bas files, got conflict:\n{lhs}\n{rhs}"
                )

    # Merge classifiers with `unknown` at their top
    for classifiers in classifiers_sources.values():
        idx = -1
        while (idx := idx + 1) < len(classifiers):
            if classifiers[idx].type == "unknown":
                for j in range(idx + 1, len(classifiers)):

                    classifiers[idx + 1].tokens = (
                        classifiers[idx].tokens + classifiers[idx + 1].tokens
                    )
                    classifiers.pop(idx)
                    if classifiers[idx].type != "unknown":
                        break

    classifiers = reduce(operator.concat, classifiers_sources.values())  # noqa

    buckets = {
        "attribute": [],
        "option": [],
        "declare": [],
        "global": [],
        "other": [],
        "function": [],
    }

    for c in classifiers:
        # Only use one of the file's declare statements
        if c.type == "option":
            if c.origin == first(classifiers_sources):
                buckets[c.type].append(c)
        elif c.type in buckets:
            buckets[c.type].append(c)
        else:
            buckets["other"].append(c)

    code = []
    origin = None
    for key, vals in buckets.items():
        if key == "attribute":
            continue

        for i in vals:
            block_str = to_unix_line_endings(tokens_to_str(i.tokens)).strip()
            if block_str != "":
                block_str = block_str + "\n\n"
                if i.type != "option":
                    if i.origin != origin:
                        code.append(f"'*************** {names[i.origin]}\n")
                        origin = i.origin

                code.append(block_str)

    if module_name is None:
        module_name = first(names.values())

    code.insert(0, f'Attribute VB_Name = "{module_name}"\n')

    return "".join(code)
