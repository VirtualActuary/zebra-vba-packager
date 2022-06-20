import re
from dataclasses import dataclass
from functools import reduce
import operator
from typing import List, Tuple
from itertools import chain

# https://github.com/rubberduck-vba/Rubberduck/issues/3175 amended with some experimental findings of our own
vba_names = """
    Abs Access AddressOf Alias And Any Append Array As Assert Attribute B
    BF Base Binary Boolean ByRef ByVal Byte CBool CByte CCur CDate CDbl
    CDec CDecl CInt CLng CLngLng CLngPtr CSng CStr CVDate CVErr CVar Call
    Case ChDir Circle Close Compare Const CurDir Currency Database Date
    Debug Decimal Declare DefBool DefByte DefCur DefDate DefDbl DefDec
    DefInt DefLng DefLngLng DefLngPtr DefObj DefSng DefStr DefVar Dim Dir
    Do DoEvents Double Each Else ElseIf Empty End EndIf Enum Eqv Erase
    Error Event Exit Explicit F False Fix For Format FreeFile Friend
    Function Get Global Go GoSub GoTo If IIf Imp Implements In InStr InStrB
    Input InputB Int Integer Is LBound LINEINPUT LSet Left Len LenB Let Lib
    Like Line Load Local Lock Long LongLong LongPtr Loop Me Mid MidB Mod
    Module MultiUse Name New Next Not Nothing Null Object On Open Option
    Optional Or Output PSet ParamArray Preserve Print Private Property
    PtrSafe Public Put RGB RSet RaiseEvent Random Randomize ReDim Read Rem
    Resume Return Scale Seek Select Set Sgn Shared Single Spc Static Step
    Stop StrComp String Sub Tab Text Then ThisWorkbook To True Type TypeOf
    UBound Unknown Unload Unlock Until VB_Base VB_Control VB_Creatable
    VB_Customizable VB_Description VB_Exposed VB_Ext_KEY VB_GlobalNameSpace
    VB_HelpID VB_Invoke_Func VB_Invoke_Property VB_Invoke_PropertyPut
    VB_Invoke_PropertyPutRef VB_MemberFlags VB_Name VB_PredeclaredId
    VB_ProcData VB_TemplateDerived VB_UserMemId VB_VarDescription
    VB_VarHelpID VB_VarMemberFlags VB_VarProcData VB_VarUserMemId Variant
    Wend While Width Win32 Win64 With WithEvents Write Workbook Xor
    """.split()

vba_names_set = set([i.lower() for i in vba_names])

linecont_re = re.compile("[\t ](_\n)")

filler = r"nt ^()&-+*/=,.[]"

filler_re_str = "^|" + ("|".join(["\\" + i for i in filler]))

attribname_re = re.compile(
    r'attribute[\t ]+vb_name[\t ]*=[\t ]*"([a-z][a-z0-9_]*)"', re.IGNORECASE
)

name_re = re.compile(rf"({filler_re_str})([a-zA-Z][a-zA-Z0-9_]*)")

hashif_re = re.compile(
    rf"({filler_re_str})((#if)|(#elseif)|(#else)|(#end[\t ].*if))", re.IGNORECASE
)

# from openpyxl/formula/tokenizer.py
# Inside a string, all characters are treated as literals, except for
# the quote character used to start the string. That character, when
# doubled is treated as a single character in the string. If an
# unmatched quote appears, the string is terminated.
string_re = re.compile('"(?:[^"]*"")*[^"]*"(?!")')

# https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rem-statement
comment_re = re.compile(r"((')|(:[ \t]*rem[ \t])).*", re.IGNORECASE)

commentrem_re = re.compile(r"[ \t]*rem[ \t].*", re.IGNORECASE)

whitespace_re = re.compile(r"([ \t]|(_\n))+")

newline_re = re.compile(r"\n")


def str_idxes(s):
    idelta = 0
    lines = s.split("\n")
    for line in lines:
        for i, j in re_idx(string_re, line):
            yield i + idelta, j + idelta
        idelta += len(line) + 1


def comment_idxes(s):
    idelta = 0
    idxes = []
    lines = s.split("\n")
    comment_cont = False
    for line in lines:
        if comment_cont:
            idxes[-1] = (idxes[-1][0], idelta + len(line))
            if not line[-1:] == "_":
                comment_cont = False
        else:
            idx = []
            if (x := re_idx(commentrem_re, line)) and x[0][0] == 0:
                idx = [(idelta + x[0][0], idelta + len(line))]
            elif x := re_idx(comment_re, line):
                idx = [(idelta + x[0][0], idelta + len(line))]

            idxes.extend(idx)

            if idx and line[-1:] == "_":
                comment_cont = True

        idelta += len(line) + 1

    return idxes


@dataclass
class VBAToken:
    text: str
    type: str  # "name" "unknown"


def prod(lst):
    return reduce(operator.mul, lst, 1)


def re_idx(reg, s, group_nr=0):
    """
    Get [i to j) index of regex expression
    """
    return [(m.start(group_nr), m.end(group_nr)) for m in re.finditer(reg, s)]


def replace_with_spaces(s: str, idxes, invert=False, spacechar=" "):
    assert len(spacechar) == 1

    sparts = []
    for (i, j), active in idx_sequencing(idxes, len(s)):
        if (invert and active) or (not invert and not active):
            sparts.append(s[i:j])
        else:
            sparts.append(spacechar * (j - i))

    return "".join(sparts)


def idx_sequencing(idxes: List[Tuple[int, int]], last):
    jprev = 0
    for i, j in idxes:
        yield (jprev, i), False
        yield (i, j), True
        jprev = j

    yield (jprev, last), False


def tokenize(txt) -> List[VBAToken]:
    r"""
    This tokenizes VBA into tokens - it only impliments the type of tokens that we need for now, the rest are given a
    type of "unknown"

    :param txt: string consisting of VBA code
    :return: VBA tokens

    >>> vba_txt = '''Attribute VB_Name = "FileCompress" 'comment
    ... Option Compare _
    ... Text
    ... Option Explicit
    ...
    ... ' Compression and decompression methods v1.1.1
    ... #If EarlyBinding = False Then
    ...     Public Enum IOMode
    ...         ForAppending = 8
    ...         ForReading = 1
    ...         ForWriting = 2
    ...     End Enum
    ... #End If
    ... '''

    >>> tokens = tokenize(vba_txt)
    >>> "".join([i.text for i in tokenize(vba_txt)]).replace("\r", "") == vba_txt
    True

    """

    idxmap = {}
    txt = txt.replace("\r\n", "\n")
    lower = txt.lower()
    s = txt

    # Hack to make xxx in 'attribute vb_name = "xxx"' a name and not a string
    if idx := re_idx(attribname_re, s, 1):
        i, j = idx[0]
        idxmap[(i, j)] = "name"
        s = s[: i - 1] + "·" * (j - i + 2) + s[j + 1 :]

    # Replace possible string entries to not contain comment starters like ' or REM
    s = replace_with_spaces(s, [[i + 1, j - 1] for i, j in str_idxes(s)])
    s = replace_with_spaces(s, comment_idx := list(comment_idxes(s)))
    idxmap.update((i, "comment") for i in comment_idx)

    s = replace_with_spaces(s, str_idx := list(str_idxes(s)))
    idxmap.update((i, "string") for i in str_idx)

    s = replace_with_spaces(s, hashif_idx := list(re_idx(hashif_re, s, 2)))
    idxmap.update((i, "#if") for i in hashif_idx)

    s = replace_with_spaces(
        s,
        wspace_idx := list(
            re_idx(
                whitespace_re, replace_with_spaces(s, sorted(idxmap), spacechar="x"), 0
            )
        ),
    )
    idxmap.update((i, "space") for i in wspace_idx)

    idxmap.update(
        (i, "name" if lower[i[0] : i[1]] not in vba_names_set else "reserved")
        for i in list(re_idx(name_re, s, 2))
    )
    idxmap.update((i, "newline") for i in list(re_idx(newline_re, s, 0)))

    tokens = []
    for (i, j), known in idx_sequencing(sorted(idxmap), len(s)):
        if i == j:
            continue
        if known:
            ttype = idxmap[(i, j)]
            if ttype == "newline":
                tokens.append(VBAToken(text="\n", type=ttype))
            else:
                tokens.append(VBAToken(text=txt[i:j], type=ttype))
        else:
            tokens.append(VBAToken(text=txt[i:j], type="unknown"))

    return tokens


def tokens_to_str(tokens: List[VBAToken]) -> str:
    return "".join([i.text for i in tokens])
