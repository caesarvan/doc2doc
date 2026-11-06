"""Assemble the two hand-off manuals: inline the shared format spec and the exact code files."""
import re
import sys
from pathlib import Path

HERE = Path(__file__).parent
ROOT = HERE.parent
OUT = ROOT / "deliver"


def render(template):
    text = (HERE / template).read_text(encoding="utf-8")
    text = re.sub(r"\{\{INCLUDE:([^}]+)\}\}", lambda m: (HERE / m.group(1)).read_text(encoding="utf-8").strip(), text)
    return re.sub(r"\{\{FILE:([^}]+)\}\}",
                  lambda m: f"````python\n{(ROOT / m.group(1)).read_text(encoding='utf-8').rstrip()}\n````", text)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    OUT.mkdir(exist_ok=True)
    for src, dst in (("A.md", "A_ppt2md_DeepSeek-opencode.md"), ("B.md", "B_md2ppt_Codex.md")):
        (OUT / dst).write_text(render(src), encoding="utf-8")
        print(OUT / dst, f"{(OUT / dst).stat().st_size / 1024:.0f} KB")
