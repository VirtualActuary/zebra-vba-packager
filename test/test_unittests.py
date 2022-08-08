from pprint import pprint
from textwrap import dedent

import locate
import unittest

with locate.prepend_sys_path(".."):
    from zebra_vba_packager.bas_combining import (
        compile_code_into_sections,
        compile_bas_sources_into_single_file,
    )
    from zebra_vba_packager.vba_tokenizer import tokens_to_str, tokenize
    from zebra_vba_packager.match_tokens import match_tokens


def lstripdedent(s):
    return dedent(s).lstrip()


class TestBasCombining(unittest.TestCase):
    def test_matching(self):
        x = "\nPrivate Function Bla2(arr As Variant)\n    Bla2 = True\nEnd Function"

        self.assertEqual(
            list(
                match_tokens(
                    tokenize(x),
                    "[private|public] property|sub|function",
                    on_line_start=True,
                )
            ),
            [(1, 4)],
        )

    def test_compile_code_into_sections_functions(self):

        x = lstripdedent(
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
            
            """
        )

        y = compile_code_into_sections(x)

        self.assertEqual(
            [i.type for i in y],
            [
                "attribute",
                "option",
                "unknown",
                "function",
                "unknown",
                "function",
                "unknown",
            ],
        )

    def test_compile_code_into_sections_enum(self):
        x = lstripdedent(
            """
            Attribute VB_Name = "aErrorEnums"
            ' Comment
            
            Option Explicit
            
            Enum ErrNr
                Val1 = 3
                Val2 = 5
            End Enum
            
            """
        )

        y = compile_code_into_sections(x)

        self.assertEqual(
            ["attribute", "unknown", "option", "unknown", "enum", "unknown"],
            [i.type for i in y],
        )

        self.assertTrue("".join([tokens_to_str(i.tokens) for i in y]) == x)

    def test_compile_code_into_sections_ptrsafe(self):
        x = lstripdedent(
            """
            Attribute VB_Name = "MiscAssign"
            
            Option Explicit
            Option Private
            
            Private Declare PtrSafe Function ShellExecuteA Lib "Shell32.dll" _
               (ByVal hwnd As Long, _
               ByVal lpOperation As String, _
               ByVal lpFile As String, _
               ByVal lpParameters As String, _
               ByVal lpDirectory As String, _
               ByVal nShowCmd As Long) As Long
               
            """
        )

        y = compile_code_into_sections(x)

        self.assertEqual(
            [
                "attribute",
                "unknown",
                "option",
                "option",
                "unknown",
                "declare",
                "unknown",
            ],
            [i.type for i in y],
        )

    def test_compile_code_into_sections_hashif(self):
        x = lstripdedent(
            """
            Attribute VB_Name = "MiscAssign"
            
            Option Explicit
            Option Private
            
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
            """
        )

        y = compile_code_into_sections(x)

        self.assertEqual(
            ["attribute", "unknown", "option", "option", "unknown", "#if", "unknown"],
            [i.type for i in y],
        )

        self.assertTrue("".join([tokens_to_str(i.tokens) for i in y]) == x)

    def test_underscore_names(self):
        txt = lstripdedent(
            """
        Dim cn As WorkbookConnection
        On Error GoTo err_
        Application.Calculation = xlCalculationManual
        Dim numConnections As Integer, i As Integer
           End If
        Next
        GoTo done_
        err_:
        done_:
        """
        )

        self.assertEqual(tokens_to_str(tokenize(txt)), txt)

    def test_combine_bas_sources_into_single_file(self):
        sources = {
            "file_a": lstripdedent(
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
                """
            ),
            "file_b": lstripdedent(
                """
                Attribute VB_Name = "aErrorEnums"
                ' Comment

                Option Explicit

                Enum ErrNr
                    Val1 = 3
                    Val2 = 5
                End Enum
                """
            ),
            "file_c": lstripdedent(
                """
                Attribute VB_Name = "MiscAssign"

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

                Private Function Bla3()
                    Bla3 = True
                End Function
                """
            ),
        }

        x = compile_bas_sources_into_single_file(sources)
        y = lstripdedent(
            """
            Attribute VB_Name = "MiscArray"
            Option Explicit
            
            '*************** MiscAssign
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
            
            '*************** aErrorEnums
            Enum ErrNr
                Val1 = 3
                Val2 = 5
            End Enum
            
            '*************** MiscArray
            ' Comment
            Private Function MiscArray_Bla(arr As Variant)
                MiscArray_Bla = True
            End Function
            
            Private Function MiscArray_Bla2(arr As Variant)
                MiscArray_Bla2 = True
            End Function
            
            '*************** MiscAssign
            Private Function MiscAssign_Bla3()
                MiscAssign_Bla3 = True
            End Function
            
            """
        )

        self.assertEqual(y, x.replace("\r", ""))

        with self.assertRaisesRegex(
            ValueError,
            "Options must be equal across aggregated bas files, got conflict.*",
        ):
            sources["file_b"] = f"Option Private\n{sources['file_b']}"
            compile_bas_sources_into_single_file(sources)


if __name__ == "__main__":
    import unittest

    unittest.main()
