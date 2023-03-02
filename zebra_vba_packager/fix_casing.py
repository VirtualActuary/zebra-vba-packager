from .vba_tokenizer import vba_names, tokenize, tokens_to_str
import os
from typing import Union
from pathlib import Path


def fix_casing(
    code_dir: Union[str, Path],
    case_style: Union[str, None] = None,
    vars_overwrite_file: Union[str, Path, list] = None,
):
    """
    Enforce VBA casing.
    This includes:
        - Checking that variables start with the correct casing
        - Checking that variables casing is consistent across all files.
        - Filenames start with the correct casing
        - DLL variables remain unchanged.
        - Excluding desired variables.

    Args:
        code_dir: List of directories containing the '.bas', '.cls', '.txt' files to enforce casing on.
        case_style: Casing style. Options -> None, 'camel', 'pascal'. Not case-sensitive
        vars_overwrite_file: File containing variable names that overwrites the selected casing option.
    """
    if case_style is not None:
        case_style = case_style.lower()
        if case_style not in ["camel", "pascal"]:
            raise ValueError("Only camel and pascal casing supported.")

    exts = [".bas", ".cls", ".txt"]
    vba_names_dict = {name.lower(): name for name in vba_names}

    tokens_dict = {}
    if vars_overwrite_file is not None:
        tokens_dict = _fetch_vars_overwrite_file_data(vars_overwrite_file)

    vbafiles = _get_vba_filenames(code_dir, exts)

    for file_name_input in vbafiles:
        file_name_output = _rename_filename(file_name_input, case_style)

        with open(file_name_input) as file:
            file_content_input = file.read()
        tokens = tokenize(file_content_input)

        tokens, tokens_dict = _convert_token_txt(
            tokens, case_style, tokens_dict, vba_names_dict
        )

        file_content_output = tokens_to_str(tokens)

        # For filename changes
        if file_name_input != file_name_output:
            os.remove(file_name_input)

        # For any change
        if (
            file_content_output != file_content_input
            or file_name_input != file_name_output
        ):
            with open(file_name_output, "w") as ofile:
                ofile.write(file_content_output)


def _rename_filename(f, case_style):
    file_name_old = os.path.basename(f)[:-4]
    file_name_new = _change_variable_casing(file_name_old, case_style)
    return (
        os.path.dirname(f)
        + "\\"
        + os.path.basename(f).replace(file_name_old, file_name_new)
    )


def _fetch_vars_overwrite_file_data(vars_overwrite_file):
    tokens_dict = {}
    try:
        with open(vars_overwrite_file) as file:
            overwrites = file.read().split("\n")
    except Exception as e:
        print(e)
    else:
        for line in overwrites:
            line = line.split("#")[0].strip()
            if line:
                tokens_dict[line.lower()] = line
    return tokens_dict


def _change_variable_casing(variable: str, case_style: str):
    if case_style is None:
        return variable

    if case_style == "camel":
        if len(variable) == 1:
            return variable.lower()

        elif not variable.upper() == variable:
            return variable[:1].lower() + variable[1:]
        return variable

    if case_style == "pascal":
        variable = variable[0].upper() + variable[1:]
        return variable


def _convert_token_txt(tokens, case_style, tokens_dict, vba_names_dict):
    for index, token in enumerate(tokens):
        if token.type == "name":  # variables
            if _token_is_dll_name(tokens, index):
                continue

            if token.text.lower() in tokens_dict.keys():
                tokens[index].text = tokens_dict[token.text.lower()]
            else:
                tokens[index].text = _change_variable_casing(token.text, case_style)
                tokens_dict[token.text.lower()] = tokens[index].text

        elif token.type == "reserved":  # reserved VBA names
            tokens[index].text = vba_names_dict[token.text.lower()]

    return tokens, tokens_dict


def _token_is_dll_name(tokens, index) -> bool:
    if index < 2:  # no space for "function" or "sub"
        return False

    # token.text:function/sub -> token:whitespace -> token:name -> token:whitespace -> token.text:lib
    if tokens[index - 2].type == "reserved" and tokens[index - 2].text.lower() in (
        "function",
        "sub",
    ):
        if (
            tokens[index + 2].type == "reserved"
            and tokens[index + 2].text.lower() == "lib"
        ):
            return True
    return False


def _get_vba_filenames(codedir, exts):
    vbafiles = []
    for d, dirs, files in os.walk(codedir):
        for f in files:
            if not f.lower()[-4:] in exts:
                continue

            path = os.path.join(d, f)
            vbafiles.append(path)
    return vbafiles
