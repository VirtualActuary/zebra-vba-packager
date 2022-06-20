from textwrap import dedent
from types import SimpleNamespace as SN
from typing import List
import re

from .vba_tokenizer import VBAToken, tokenize


def _str_to_matchables(s):
    matchlist = [i.strip() for i in s.split(" ")]
    matchobj = []
    for i in matchlist:
        obj = SN(optional=False, re=i)
        if f"{i[:1]}{i[-1:]}" == "[]":  # optional
            obj.optional = True
            obj.re = i[1:-1]

        obj.re = re.compile(obj.re, flags=re.IGNORECASE)

        matchobj.append(obj)
    return matchobj


def match_tokens(
    tokens: List[VBAToken], custom_token_match_string, on_line_start=False
):
    matchables = _str_to_matchables(custom_token_match_string)

    # Pretend that tokens[-1] is a newline
    def token_at(i):
        return tokens[i] if i >= 0 else VBAToken(type="newline", text="\n")

    # For on_line_start start at tokens[-1] and inject first match as =="\r\n"
    if on_line_start:
        matchables.insert(0, SN(optional=False, re=SN(match=lambda x: x == "\n")))
        i = -2
    else:
        i = -1

    while (i := i + 1) < len(tokens):
        if token_at(i).type == "space":
            continue

        matched = 1  # 1 = ongoing
        k = -1
        j = i - 1
        while (j := j + 1) < len(tokens) and k < len(matchables):
            if token_at(j).type == "space":
                continue

            while (k := k + 1) < len(matchables):

                if matchables[k].re.match(token_at(j).text):
                    if k == len(matchables) - 1:
                        matched = 2  # final match
                    break

                elif matchables[k].optional:  # Optional don't have to match
                    continue
                else:
                    matched = 0  # early no match termination
                    break

            if matched != 1:
                break

        if matched == 2:
            ii = i
            while on_line_start and tokens[(ii := ii + 1)].type == "space":
                pass

            yield ii, j + 1


if __name__ == "__main__":
    txt_in = dedent(
        """
        Attribute VB_Name = "MiscArray"
        Option Explicit
        
        ' Comment
        Private Function Bla(arr As Variant)
            Bla = True
        End Function
        Private Function Bla2(arr As Variant)
            Bla2 = True
        End Function
        
        
        Enum ErrNr
            Val1 = 3
            Val2 = 5
        End Enum
        
        
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
                
            Declare Function ShellExecuteA Lib "Shell32.dll"
        #End If
        
        Private Function Bla3()
            Bla3 = True
        End Function
        """
    )

    tokens = tokenize(txt_in)
    (i, j) = next(match_tokens(tokens, "#if", on_line_start=True))

    print(tokenize("hello\n  #end if"))
    print(
        list(
            match_tokens(tokenize("hello\n  #end if"), "#end.*if", on_line_start=False)
        )
    )
