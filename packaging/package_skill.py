#!/usr/bin/env python3
"""Zip the Claude skill for distribution.

    python packaging/package_skill.py                              # dist/markitdown-skill.zip
    python packaging/package_skill.py --exe dist/markitdown.exe    # + dist/markitdown-skill-windows.zip

Both archives contain a top-level `markitdown/` folder with SKILL.md. Unzip it into
~/.claude/skills/ (Claude Code) or upload the zip in the Skills section of the
Claude app settings. The Windows variant also ships bin/markitdown.exe so the
skill works without Python.
"""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILL_DIR = ROOT / "skills" / "markitdown"


def skill_files():
    for path in sorted(SKILL_DIR.rglob("*")):
        if path.is_dir() or "__pycache__" in path.parts or path.suffix == ".pyc":
            continue
        if path.relative_to(SKILL_DIR).parts[0] == "bin":  # local exe copies
            continue
        yield path


def write_zip(target: Path, exe: Path | None = None) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for path in skill_files():
            z.write(path, Path("markitdown") / path.relative_to(SKILL_DIR))
        if exe is not None:
            z.write(exe, Path("markitdown") / "bin" / exe.name)
    print(f"wrote {target} ({target.stat().st_size / 1_048_576:.1f} MB)")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--exe", help="path to markitdown.exe to bundle into a Windows variant")
    parser.add_argument("--dist", default=str(ROOT / "dist"), help="output folder (default: dist/)")
    args = parser.parse_args()
    dist = Path(args.dist)

    write_zip(dist / "markitdown-skill.zip")
    if args.exe:
        write_zip(dist / "markitdown-skill-windows.zip", Path(args.exe))


if __name__ == "__main__":
    main()
