from dataclasses import dataclass
from pprint import pprint
from textwrap import indent
from typing import Dict, Union, List
from functools import reduce
import operator

from vba_tokenizer import VBAToken, tokenize, tokens_to_str


@dataclass
class VBASectionClassifier:
    tokens: List[VBAToken]
    type: Union[str, None] = None
    origin: Union[str, None] = None
    name: Union[str, None] = None
    private: Union[bool, None] = None

    def __post_init__(self):
        if self.type is None:
            self.type = "unknown"


def compile_code_into_sections(
        input: Union[List[VBAToken], str],
        origin: Union[str, None] = None
) -> List[VBASectionClassifier]:
    r"""
    Split VBA code into a few high-level catagories, such as `#if`, `function`, `option`, `declare`, `enum`, and
    `unknown`. Extend `unknown` into more catagories as they are needed by other parts of the codebase.

    Examples:
        >>> x = '''Attribute VB_Name = "MiscArray"
        ...
        ... Option Explicit
        ...
        ... ' Comment
        ... Private Function Bla(arr As Variant)
        ...     Bla = True
        ... End Function
        ... Private Function Bla2(arr As Variant)
        ...     Bla2 = True
        ... End Function
        ... '''
        >>> y = compile_code_into_sections(x)

        >>> [i.type for i in y]
        ['attribute', 'unknown', 'option', 'unknown', 'function', 'unknown', 'function', 'unknown']

        >>> "".join([tokens_to_str(i.tokens) for i in y]).replace('\r\n', '\n') == x
        True

        >>> x = '''Attribute VB_Name = "aErrorEnums"
        ... ' Comment
        ...
        ... Option Explicit
        ...
        ... Enum ErrNr
        ...     Val1 = 3
        ...     Val2 = 5
        ... End Enum
        ... '''
        >>> y = compile_code_into_sections(x)

        >>> [i.type for i in y]
        ['attribute', 'unknown', 'option', 'unknown', 'enum', 'unknown']

        >>> "".join([tokens_to_str(i.tokens) for i in y]).replace('\r\n', '\n') == x
        True

        >>> x = '''Attribute VB_Name = "MiscAssign"
        ...
        ... Option Explicit
        ... Option Private
        ...
        ... Private Declare PtrSafe Function ShellExecuteA Lib "Shell32.dll" _
        ...    (ByVal hwnd As Long, _
        ...    ByVal lpOperation As String, _
        ...    ByVal lpFile As String, _
        ...    ByVal lpParameters As String, _
        ...    ByVal lpDirectory As String, _
        ...    ByVal nShowCmd As Long) As Long
        ... '''
        >>> y = compile_code_into_sections(x)

        >>> [i.type for i in y]
        ['attribute', 'unknown', 'option', 'unknown', 'option', 'unknown', 'declare', 'unknown']


        >>> x = '''Attribute VB_Name = "MiscAssign"
        ...
        ... Option Explicit
        ... Option Private
        ...
        ... #If VBA7 And Win64 Then
        ...     Private Declare PtrSafe Function ShellExecuteA Lib "Shell32.dll" _
        ...         (ByVal hwnd As Long, _
        ...         ByVal lpOperation As String, _
        ...         ByVal lpFile As String, _
        ...        ByVal lpParameters As String, _
        ...         ByVal lpDirectory As String, _
        ...         ByVal nShowCmd As Long) As Long
        ... #Else
        ...
        ...     Private Declare Function ShellExecuteA Lib "Shell32.dll" _
        ...         (ByVal hwnd As Long, _
        ...         ByVal lpOperation As String, _
        ...         ByVal lpFile As String, _
        ...         ByVal lpParameters As String, _
        ...         ByVal lpDirectory As String, _
        ...         ByVal nShowCmd As Long) As Long
        ... #End If
        ... '''
        >>> y = compile_code_into_sections(x)

        >>> [i.type for i in y]
        ['attribute', 'unknown', 'option', 'unknown', 'option', 'unknown', '#if', 'unknown']

        >>> "".join([tokens_to_str(i.tokens) for i in y]).replace('\r\n', '\n') == x
        True


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
            origin=classifiers[idx].origin
        )
        c3 = VBASectionClassifier(
            tokens[end_marker:],
            origin=classifiers[idx].origin
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


    # Find #if
    idx = -1
    while(idx := idx+1) < len(classifiers):
        if classifiers[idx].type != "unknown":
            continue

        tokens = classifiers[idx].tokens

        # Merge spaces together
        i = -1
        while (i := i + 1) < len(tokens):
            start_marker = i
            end_marker = None
            if tokens[i].type == "#if" and tokens[i].text.lower() == "#if":
                j = i
                while (j := j + 1) < len(tokens):
                    if tokens[j].type == "#if" and tokens[j].text.lower().startswith("#end"):
                        end_marker = j+1
                        break

            if end_marker is not None:
                extend_classification_list(
                    classifiers,
                    idx,
                    start_marker,
                    end_marker,
                    params2=VBASectionClassifier([], "#if")
                )
                idx -= 1
                break

    # Find function and enum
    func_lables = {'property', 'sub', 'function', 'enum'}
    idx = -1
    while(idx := idx+1) < len(classifiers):
        if classifiers[idx].type != "unknown":
            continue

        tokens = classifiers[idx].tokens

        # Find sandwich for `[private/public] Function .... End Function`
        i = len(tokens)
        while (i := i - 1) >= 0:
            if (tokens[i].type == 'reserved' and tokens[i].text.lower() in func_lables
                    and i-1 >= 0 and tokens[i-1].type == 'space'
                    and i-2 >= 0 and tokens[i-2].text.lower() == 'end'):
                params2 = VBASectionClassifier([], 'function')
                if tokens[i].text.lower() == 'enum':
                    params2.type = 'enum'

                start_marker = None
                end_marker = i+1

                j = i-2
                while (j := j - 1) >= 0:
                    if (tokens[j].type == 'reserved' and
                            tokens[j].text.lower() in func_lables):

                        #allowed `[<private/public>]<space>function`
                        if (j - 2 >= 0 and tokens[j - 1].type == "space"
                                and tokens[j - 2].text.lower() in ('private', 'public')):
                            start_marker = j - 2
                        else:
                            start_marker = j
                        break

                if start_marker is not None:
                    extend_classification_list(classifiers, idx, start_marker, end_marker, params2=params2)
                    idx -= 1
                    break

    # Find option
    idx = -1
    while(idx := idx+1) < len(classifiers):
        if classifiers[idx].type != "unknown":
            continue

        tokens = classifiers[idx].tokens

        i = -1
        while (i := i + 1) < len(tokens):
            start_marker = i
            end_marker = None
            params2 = None

            if tokens[i].type == "reserved":
                ttext = tokens[i].text.lower()

                if ttext == 'declare':
                    params2 = 'declare'
                    end_marker = len(tokens)

                elif ttext in ('private', 'public'):
                    if (i + 2 >= 0 and tokens[i + 1].type == "space"
                            and tokens[i + 2].text.lower() == 'declare'):
                        params2 = 'declare'
                        end_marker = len(tokens)

                elif ttext == 'option':
                    params2 = 'option'
                    end_marker = len(tokens)

                elif ttext == 'attribute':
                    params2 = 'attribute'
                    end_marker = len(tokens)


            if end_marker is not None:
                for j in range(i + 1, len(tokens)):
                    if tokens[j].type == 'newline':
                        end_marker = j
                        break

                extend_classification_list(classifiers, idx, start_marker, end_marker, params2=params2)
                idx -= 1
                break

    # Remove empty classifications
    idx = -1
    while(idx := idx+1) < len(classifiers):
        if not len(classifiers[idx].tokens):
            classifiers.pop(idx)
            idx -= 1

    return classifiers


def compile_bas_sources_into_single_file(
        sources: Dict[str, Union[str, List[VBAToken]]],
        module_name: Union[str, None] = None
) -> List[VBASectionClassifier]:
    r"""
    Examples:
        >>> sources = {
        ... 'file_a': '''Attribute VB_Name = "MiscArray"
        ...
        ... Option Explicit
        ...
        ... ' Comment
        ... Private Function Bla(arr As Variant)
        ...     Bla = True
        ... End Function
        ... Private Function Bla2(arr As Variant)
        ...     Bla2 = True
        ... End Function
        ... ''',
        ... 'file_b': '''Attribute VB_Name = "aErrorEnums"
        ... ' Comment
        ...
        ... Option Explicit
        ...
        ... Enum ErrNr
        ...     Val1 = 3
        ...     Val2 = 5
        ... End Enum
        ... ''',
        ... 'file_c': '''Attribute VB_Name = "MiscAssign"
        ...
        ... Option Explicit
        ...
        ... #If VBA7 And Win64 Then
        ...     Private Declare PtrSafe Function ShellExecuteA Lib "Shell32.dll" _
        ...         (ByVal hwnd As Long, _
        ...         ByVal lpOperation As String, _
        ...         ByVal lpFile As String, _
        ...        ByVal lpParameters As String, _
        ...         ByVal lpDirectory As String, _
        ...         ByVal nShowCmd As Long) As Long
        ... #Else
        ...
        ...     Private Declare Function ShellExecuteA Lib "Shell32.dll" _
        ...         (ByVal hwnd As Long, _
        ...         ByVal lpOperation As String, _
        ...         ByVal lpFile As String, _
        ...         ByVal lpParameters As String, _
        ...         ByVal lpDirectory As String, _
        ...         ByVal nShowCmd As Long) As Long
        ... #End If
        ... '''
        ... }
        >>> _ = compile_bas_sources_into_single_file(sources)

        >>> sources['file_b'] =  'Option Private' + '\n' + sources['file_b']
        >>> _ = compile_bas_sources_into_single_file(sources) # doctest: +IGNORE_EXCEPTION_DETAIL
        Traceback (most recent call last):
        ValueError: Options must be equal across aggregated bas files, got conflict:
         ...

    """

    sources = {key: val if isinstance(val, list) else tokenize(val) for key, val in sources.items()}

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
                normcode = ("".join([' ' if i.type == 'space' else i.text for i in c.tokens]).lower()).strip()
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

                    classifiers[idx+1].tokens = classifiers[idx+1].tokens + classifiers[idx+1].tokens
                    classifiers.pop(idx)
                    if classifiers[idx].type != 'unknown':
                        break

    classifiers = reduce(operator.concat, classifiers_sources.values())

    buckets = {
        'attribute': [],
        'option': [],
        'declare': [],
        'enum': [],
        'other': []
    }

    for c in classifiers:
        if c.type in buckets:
            buckets[c.type].append(c)
        else:
            buckets['other'].append(c)

    code = []
    origin = None
    for keys, vals in buckets.items():
        if key == 'attribute':
            continue

        for i in vals:
            block_str = tokens_to_str(i.tokens).replace("\r", "").strip().replace("\n", "\r\n")
            if block_str != "":
                if i.origin != origin:
                    code.append(f"*************** {i.origin}\r\n")
                    origin = i.origin

                code.append(block_str)
                code.append("\r\n\r\n")

    if module_name is not None:
        code.insert(0, f'Attribute VB_Name = "{module_name}"\r\n')

    return "".join(code)


if __name__ == "__main__":
    sources = {
        'file_a': '''Attribute VB_Name = "MiscArray"

     Option Explicit

     ' Comment
     Private Function Bla(arr As Variant)
         Bla = True
     End Function
     Private Function Bla2(arr As Variant)
         Bla2 = True
     End Function
     ''',
        'file_b': '''Attribute VB_Name = "aErrorEnums"
     ' Comment

     Option Explicit

     Enum ErrNr
         Val1 = 3
         Val2 = 5
     End Enum
     ''',
        'file_c': '''Attribute VB_Name = "MiscAssign"

     Option Explicit

     #If VBA7 And Win64 Then
         Private Declare PtrSafe Function ShellExecuteA Lib "Shell32.dll" _
             (ByVal hwnd As Long, _
             ByVal lpOperation As String, _
             ByVal lpFile As String, _
            ByVal lpParameters As String, _
             ByVal lpDirectory As String, _
             ByVal nShowCmd As Long) As Long
     #Else

         Private Declare Function ShellExecuteA Lib "Shell32.dll" _
             (ByVal hwnd As Long, _
             ByVal lpOperation As String, _
             ByVal lpFile As String, _
             ByVal lpParameters As String, _
             ByVal lpDirectory As String, _
             ByVal nShowCmd As Long) As Long
     #End If
     '''
    }

    compile_bas_sources_into_single_file(sources)