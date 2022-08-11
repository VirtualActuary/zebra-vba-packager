from __future__ import annotations
import dataclasses
import re
from pathlib import Path
from typing import Optional

from .util import read_txt, write_txt

re_modname = re.compile(
    r'^\s*Attribute\s*VB_Name\s*=\s*"(.*)"', re.IGNORECASE | re.MULTILINE
)
re_zebra_namespace = re.compile(
    r"^'\s*zebra\s*NameSpace\s*(.*)\s*$", re.IGNORECASE | re.MULTILINE
)


def fix_module_name_length_limitation(dirpath):
    """ """
    modname_reuse = set()
    for p in sorted(Path(dirpath).rglob("*.cls")):
        txt = read_txt(p)

        i, j = re_modname.search(txt).span(1)
        modname = txt[i:j]

        if modname.startswith("z__"):
            mpair = _ModnamePair.from_str(txt, suffix=1)

            if mpair.needs_renaming():
                while mpair.modname in modname_reuse:
                    mpair.suffix += 1
                modname_reuse.add(mpair.modname)
                write_txt(p, mpair.inject_into(txt))



@dataclasses.dataclass
class _ModnamePair:
    """
    Class to help annotate z__ files in order to truncate long module names, but also preserve the original
    non-truncated name.
    """

    input_modname: Optional[str]
    input_namespace: Optional[str] = None
    suffix: Optional[int] = None

    @classmethod
    def from_str(cls, txt, suffix: Optional[int] = None):
        i, j = re_modname.search(txt).span(1)
        modname = txt[i:j]

        namespace = None
        if zns := re_zebra_namespace.search(txt):
            i, j = zns.span(1)
            namespace = txt[i:j]

        return cls(modname, namespace, suffix=suffix)

    @property
    def modname(self) -> str:
        if self.needs_renaming():
            return f"{self.input_modname[:28]}{self.str_suffix}"

    @property
    def namespace(self):
        namespace = self.input_namespace

        # Namespace derived for the first time
        if namespace is None or not namespace.startswith(
            self.input_modname[:28].lstrip("z__")
        ):
            namespace = self.input_modname.lstrip("z__")

        return namespace

    @property
    def str_suffix(self) -> str:
        if self.suffix is not None:
            return "%03d" % self.suffix
        return "xxx"

    def needs_renaming(self) -> bool:
        if len(self.input_modname) > 31:
            return True
        elif (
            len(self.input_modname) == 31
            and set(self.input_modname[-3:]).difference("0123456789") == set()
        ):
            return True

        return False

    def inject_into(self, txt: str):
        # Remove any annotations
        if zns := re_zebra_namespace.search(txt):
            i, j = zns.span()
            txt = f"{txt[:i]}{txt[j+1:]}"

        # Add annotation if modname and namespace differ
        if not self.modname.lstrip("z__") == self.namespace:
            txt_lines = txt.split("\n")
            attrend = [
                i for i, s in enumerate(txt_lines) if s.lower().startswith("attribute ")
            ][-1]
            txt = "\n".join(
                txt_lines[: attrend + 1]
                + [f"'zebra NameSpace {self.namespace}"]
                + txt_lines[attrend + 1 :]
            )

        # Add the correct modname
        i, j = re_modname.search(txt).span(1)
        txt = f"{txt[:i]}{self.modname}{txt[j:]}"

        return txt
