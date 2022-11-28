import re
import tempfile
from pprint import pprint
from textwrap import dedent

import locate
import unittest
from pathlib import Path

with locate.prepend_sys_path(".."):
    from zebra_vba_packager.bas_combining import (
        compile_code_into_sections,
        compile_bas_sources_into_single_file,
    )
    from zebra_vba_packager.vba_tokenizer import tokens_to_str, tokenize
    from zebra_vba_packager.match_tokens import match_tokens
    from zebra_vba_packager.util import to_unix_line_endings
    from zebra_vba_packager import Source, Config


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

    def test_combining_bas_sources_into_single_file2(self):
        txt1 = to_unix_line_endings(
            locate.this_dir()
            .joinpath("misc-vba-example/2022-08-08-output/combined.bas")
            .read_text()
        )
        txt2 = to_unix_line_endings(
            compile_bas_sources_into_single_file(
                {
                    i: i.read_text()
                    for i in locate.this_dir()
                    .joinpath("misc-vba-example/2022-08-08-input")
                    .glob("*")
                }
            )
        )

        self.assertEqual(txt1, txt2)


class TestFullRun(unittest.TestCase):
    def test_github_download_and_combine(self):
        Config(
            Source(
                git_source="https://github.com/VirtualActuary/MiscVBAFunctions.git",
                git_rev="8e5e8f3",
                glob_include=[
                    "MiscVBAFunctions/**/*.bas",
                    "MiscVBAFunctions/**/*.cls",
                    "**/thisworkbook.txt",
                ],
                glob_exclude=["**/Test__*"],
                combine_bas_files="Fn",
                auto_bas_namespace=True,
                auto_cls_rename=False,
            )
        ).run(git_dir := Path(locate.this_dir(), "temporary_output/misc-vba-git"))

        tmp_dir = Path(locate.this_dir(), "misc-vba-git-comparison")

        self.assertEqual(
            Path(git_dir, "z__Fn.cls").read_text(),
            Path(tmp_dir, "z__Fn.cls").read_text(),
        )

    def test_github_history(self):
        git_source = locate.this_dir().joinpath("misc-vba-git-history-example")
        git_ref = "eac3bbac2faa5b40db766d439ebac06d0638f1c1"
        git_add_version_comment_params = [None, True, False]  # None defaults to True

        for git_add_version_comment in git_add_version_comment_params:
            with tempfile.TemporaryDirectory() as tmpdir:
                Config(
                    Source(
                        git_source=str(git_source),
                        git_rev=git_ref,
                        git_add_version_comment=git_add_version_comment,
                        glob_include=["*same_source/*.bas"],
                        combine_bas_files="XXX",
                        auto_bas_namespace=True,
                    )
                ).run(tmpdir)

                zebra_lines = [
                    i
                    for i in list(Path(tmpdir).rglob("*.cls"))[0]
                    .read_text()
                    .splitlines()
                    if i.startswith("'zebra")
                ]

                if git_add_version_comment == False:
                    self.assertEqual(
                        [
                            git_add_version_comment,
                            "'zebra source i_am_from_another_repo#1234567",
                            "'zebra source i_am_from_another_repo#1234567",
                        ],
                        [git_add_version_comment] + zebra_lines,
                    )
                else:
                    self.assertEqual(
                        [
                            git_add_version_comment,
                            "'zebra source misc-vba-git-history-example#eac3bba <- i_am_from_another_repo#1234567",
                        ],
                        [git_add_version_comment] + zebra_lines,
                    )

            with tempfile.TemporaryDirectory() as tmpdir:
                Config(
                    Source(
                        git_source=str(git_source),
                        git_rev=git_ref,
                        git_add_version_comment=git_add_version_comment,
                        glob_include=["*different_source/*.bas"],
                        combine_bas_files="XXX",
                        auto_bas_namespace=True,
                    )
                ).run(tmpdir)

                zebra_lines = [
                    i
                    for i in list(Path(tmpdir).rglob("*.cls"))[0]
                    .read_text()
                    .splitlines()
                    if i.startswith("'zebra")
                ]

                if git_add_version_comment == False:
                    self.assertEqual(
                        [
                            git_add_version_comment,
                            "'zebra source i_am_from_another_repo#1234567",
                            "'zebra source i_am_from_yet_another_repo#1234567",
                        ],
                        [git_add_version_comment] + zebra_lines,
                    )
                else:
                    self.assertEqual(
                        [
                            git_add_version_comment,
                            "'zebra source misc-vba-git-history-example#eac3bba <- ...",
                        ],
                        [git_add_version_comment] + zebra_lines,
                    )


if __name__ == "__main__":
    import unittest

    unittest.main()
