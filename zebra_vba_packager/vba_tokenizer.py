import re
from dataclasses import dataclass
from functools import reduce
import operator
from typing import List

@dataclass
class VBAToken:
    text: str
    type: str #"name" "unknown"

# from openpyxl/formula/tokenizer.py
# Inside a string, all characters are treated as literals, except for
# the quote character used to start the string. That character, when
# doubled is treated as a single character in the string. If an
# unmatched quote appears, the string is terminated.
strre = re.compile('"(?:[^"]*"")*[^"]*"(?!")')
namere = re.compile(r"([a-zA-Z0-9_][a-zA-Z0-9_]*)")

# inline rem https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rem-statement
remcommentre = re.compile(r":[ \t]*[Rr][Ee][Mm][ \t].*")
remcommentsolre = re.compile(r"[ \t]*[Rr][Ee][Mm][ \t].*")

wspacere = re.compile(r"[ \t]*")

# from https://github.com/rubberduck-vba/Rubberduck/issues/3175
# Amended with some experimental findings of our own
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


def prod(lst):
    return reduce(operator.mul, lst, 1)


def re_idx(reg, s):
    """
    Get [i to j) index of regex expression
    """
    return [(m.start(0), m.end(0)) for m in reg.finditer(s)]


def names_idx(s):
    """
    Get [i to j) of names in regex expression: a to Z with numbers and _
    excluding number nad _ as starting char
    """
    # remove names that starts with a number
    idxes = [(i, j) for i, j in re_idx(namere, s) if not s[i] in "0123456789_"]
    return idxes


def fill_space(s, i, j):
    """
    Replace [i to j) slice in a string with empty space " "*len
    """
    return s[:i] + " " * (j - i) + s[j:]


def idx_splitting(idxes, length):
    flat_idxes = [outer for inner in idxes for outer in inner]
    rets = []
    for cnt, (i, j) in enumerate(zip([0]+flat_idxes, flat_idxes+[length])):
        rets.append(((i, j), cnt % 2 != 0))

    return rets


def tokenize(txt) -> List[VBAToken]:
    r"""
    This tokenizes VBA into tokens - it only impliments the type of tokens that we need for now, the rest are given a
    type of "unknown"

    :param txt: string consisting of VBA code
    :return: VBA tokens

    >>> vba_txt = '''
    ... Attribute VB_Name = "FileCompress" 'comment
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

    >>> "".join([i.text for i in tokenize(vba_txt)]).replace("\r", "") == vba_txt
    """

    tokens = []
    lines = txt.split("\n")

    for lidx, line in enumerate(txt.split("\n")):
        line_tokens = []

        # Find string candidates
        for (i, j), is_str in idx_splitting(re_idx(strre, line), len(line)):
            line_tokens.append(VBAToken(text=line[i:j],
                                        type=("string" if is_str else "unknown")))

        # Find comment candidates (and retrofix "strings" within comments)
        for i in range(len(line_tokens)):
            if line_tokens[i].type == "unknown":
                ttxt = line_tokens[i].text
        
                # Find rem-type comment
                if (remidx := re_idx(remcommentsolre, ttxt)) and len(remidx) == 1 and remidx[0][0] == [0]:
                    remidx = remidx
                elif (remidx := re_idx(remcommentre, ttxt)) and len(remidx) == 1:
                    remidx = remidx
                else:
                    remidx = None
        
                # ren-type comments
                if remidx is not None:
                    il, jl = 0, remidx[0][0]
                    ir, jr = remidx[0][0], len(ttxt)

                    lhs = ttxt[il: jl]
                    rhs = ttxt[ir: jr] + "".join([i.text for i in line_tokens[i+1:]])
        
                    line_tokens = line_tokens[:i]
                    line_tokens.append(VBAToken(text=lhs, type="unknown"))
                    line_tokens.append(VBAToken(text=rhs, type="comment"))
                    break
        
                # '-type comments
                elif "'" in ttxt:
                    lhs = ttxt[:ttxt.find("'")]
                    rhs = ttxt[ttxt.find("'"):] + "".join([i.text for i in line_tokens[i+1:]])

                    line_tokens = line_tokens[:i]
                    line_tokens.append(VBAToken(text=lhs, type="unknown"))
                    line_tokens.append(VBAToken(text=rhs, type="comment"))
                    break

        # find names and (reserved keywords)
        i = -1
        while (i := i+1) < len(line_tokens):
            if line_tokens[i].type == "unknown":
                t = line_tokens.pop(i)
                j = -1
                for (ii, jj), is_name in idx_splitting(names_idx(t.text), len(t.text)):
                    j += 1
                    ttxt = t.text[ii:jj]
                    if is_name:
                        if ttxt.lower() not in vba_names_set:
                            ttype = "name"
                        else:
                            ttype = "reserved"
                    else:
                        ttype = "unknown"

                    line_tokens.insert(
                        i+j,
                        VBAToken(text=ttxt,
                                 type=ttype)
                    )
                i = i+j

        tokens.extend(line_tokens)
        tokens.append(VBAToken(text="\r\n",  type="newline"))
    tokens = tokens[:-1] # remove last newline

    # Whitespace tokens
    i = -1
    while (i := i+1) < len(tokens):
        if tokens[i].type == "unknown":
            t = tokens.pop(i)
            j = -1
            for (ii, jj), is_wspace in idx_splitting(re_idx(wspacere, t.text), length=len(t.text)):
                j += 1
                tokens.insert(i+j, VBAToken(text=t.text[ii:jj],
                                            type="space" if is_wspace else "unknown"))
            i = i+j

    # Line continuation
    i = -1
    while (i := i+1) < len(tokens):
        if tokens[i].type == "unknown":
            if tokens[i].text.endswith("_"):
                j = i
                while (j := j + 1) < len(tokens):
                    if tokens[j].type == "newline":
                        break

                if prod([i.type == "space" for i in tokens[i+1:j-1]]):
                    t = tokens.pop(i)
                    tokens.insert(i, VBAToken(text=t.text.rsplit("_", 1)[0],
                                              type=t.type))

                    ttxt = "_"+t.text.rsplit("_", 1)[1]
                    for t in tokens[i+1:j+1]:
                        tokens.pop(i+1)
                        ttxt = ttxt + t.text

                    tokens.insert(i+1, VBAToken(text=ttxt,
                                                type="space"))

    # Filter out empty tokens
    tokens = [i for i in tokens if i.text != ""]

    # Merge spaces together
    i = -1
    while (i := i+1) < len(tokens):
        if tokens[i].type == "space" and i+1 < len(tokens) and tokens[i+1].type == "space":
            tokens[i].text = tokens[i].text + tokens[i+1].text
            tokens.pop(i+1)
            i = i-1

    # Fix `Attribute VB_Name = "..."` to allow ... to be used as a name
    i = -1
    while (i := i+1) < len(tokens):
        t = tokens[i]
        if (i+2 < len(tokens) and
            t.type == "reserved" and t.text.lower() == "attribute" and
                tokens[i+1].type == "space" and
                tokens[i+2].type == "reserved" and tokens[i+2].text.lower() == "vb_name"):

            is_vbname = False
            j = i+2

            # lookforward for vb_name value as a string
            while (j := j+1) < len(tokens) and tokens[j].type != "newline":
                if tokens[j].type == "string":
                    mid = tokens[j].text[1:-1]
                    tokens.pop(j)
                    tokens.insert(j + 0, VBAToken(text='"', type="unknown"))
                    tokens.insert(j + 1, VBAToken(text=mid, type="name"))
                    tokens.insert(j + 2, VBAToken(text='"', type="unknown"))
                    break

    return tokens
