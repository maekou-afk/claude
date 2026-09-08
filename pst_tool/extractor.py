"""Extract Outlook .pst files into a tree of .eml files using ``readpst``.

``readpst`` ships with the libpst project and is packaged as ``pst-utils``
on Debian/Ubuntu (``apt install pst-utils``) and ``libpst`` on Homebrew
(``brew install libpst``). It understands both the legacy ANSI PST format
and the modern Unicode OST/PST format used by Outlook 2003+.
"""

from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path

INSTALL_HINT = (
    "readpst was not found on PATH.\n"
    "Install it with one of:\n"
    "  Debian/Ubuntu: sudo apt-get install pst-utils\n"
    "  Fedora/RHEL:   sudo dnf install libpst\n"
    "  macOS:         brew install libpst\n"
)


class ReadPstNotFoundError(RuntimeError):
    pass


class PstExtractionError(RuntimeError):
    pass


@dataclass(frozen=True)
class ExtractionResult:
    output_dir: Path
    stdout: str
    stderr: str


def find_readpst() -> str:
    path = shutil.which("readpst")
    if not path:
        raise ReadPstNotFoundError(INSTALL_HINT)
    return path


def extract_pst(
    pst_path: Path,
    output_dir: Path,
    *,
    include_deleted: bool = False,
    jobs: int | None = None,
    overwrite: bool = True,
    utf8: bool = True,
) -> ExtractionResult:
    """Convert ``pst_path`` into a directory tree of .eml files.

    The resulting layout mirrors the PST's folder structure: each Outlook
    folder becomes a directory, and each message becomes a numbered
    ``.eml`` file inside it (readpst's "separate, recursive, eml" mode:
    ``-S -r -e``). Attachments are embedded as MIME parts within each
    .eml file, so no separate attachment extraction pass is needed.
    """
    pst_path = Path(pst_path)
    if not pst_path.is_file():
        raise FileNotFoundError(f"PST file not found: {pst_path}")

    readpst = find_readpst()
    output_dir.mkdir(parents=True, exist_ok=True)

    cmd = [readpst, "-S", "-r", "-e", "-q", "-o", str(output_dir)]
    if include_deleted:
        cmd.append("-D")
    if jobs and jobs > 1:
        cmd.extend(["-j", str(jobs)])
    if overwrite:
        cmd.append("-w")
    if utf8:
        cmd.append("-8")
    cmd.append(str(pst_path))

    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise PstExtractionError(
            f"readpst failed (exit code {proc.returncode}):\n{proc.stderr}"
        )
    return ExtractionResult(output_dir=output_dir, stdout=proc.stdout, stderr=proc.stderr)
