from __future__ import annotations
import re
from pathlib import Path

from .util import read_txt, write_txt

re_modname = re.compile(r"^'zebra source (.*)$", re.IGNORECASE | re.MULTILINE)
re_full_commit = re.compile(r"^[0-9a-f]{40}$")


def get_zebra_refs(txt):
    if matches := re_modname.findall(txt):
        # If there are multiple different refs, note such a merge with "..."
        if len(set(matches)) > 1:
            return ["..."]
        else:
            return [i.strip() for i in matches[0].strip().split(" <- ")]
    return []


def replace_zebra_refs(txt, sourcelist):
    old_sourcelist = get_zebra_refs(txt)
    if old_sourcelist != sourcelist:
        txt_lines = txt.split("\n")
        txt_lines = [i for i in txt_lines if not re_modname.match(i)]  # remove old refs

        attrend = [
            i for i, s in enumerate(txt_lines) if s.lower().startswith("attribute ")
        ][-1]
        txt = "\n".join(
            txt_lines[: attrend + 1]
            + ["'zebra source " + " <- ".join(sourcelist)]
            + txt_lines[attrend + 1 :]
        )
    return txt


def expand_zebra_refs(txt, new_repo, new_ref):
    sourcelist = get_zebra_refs(txt)
    new_item = (
        Path(new_repo).name.rsplit(".git", 1)[0]
        + "#"
        + (new_ref[:7] if re_full_commit.match(new_ref) else new_ref)
    )

    # Expand sourcelist to keep a chain of repo links
    if sourcelist and (new_item.rsplit("#")[0] == sourcelist[0].rsplit("#")[0]):
        sourcelist = [new_item] + sourcelist[1:]
    else:
        sourcelist = [new_item] + sourcelist

    return sourcelist


def add_repo_history_comment(dirpath, new_repo, new_ref):
    for p in sorted(Path(dirpath).rglob("*.cls")) + sorted(
        Path(dirpath).rglob("*.bas")
    ):
        txt = read_txt(p)
        sourcelist = expand_zebra_refs(txt, new_repo, new_ref)
        write_txt(p, replace_zebra_refs(txt, sourcelist))


def fix_repo_history_comment(dirpath):
    for p in sorted(Path(dirpath).rglob("*.cls")) + sorted(
        Path(dirpath).rglob("*.bas")
    ):
        txt = read_txt(p)
        sourcelist = get_zebra_refs(txt)
        write_txt(p, replace_zebra_refs(txt, sourcelist))
