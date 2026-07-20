from __future__ import annotations

import hashlib
import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
REQUIREMENTS = ROOT / "requirements.txt"
STAMP = VENV / ".requirements.sha256"


def venv_python() -> Path:
    if os.name == "nt":
        return VENV / "Scripts" / "python.exe"
    return VENV / "bin" / "python"


def run(command: list[str]) -> None:
    subprocess.check_call(command, cwd=ROOT)


def call(command: list[str]) -> int:
    return subprocess.call(command, cwd=ROOT)


def requirements_hash() -> str:
    if not REQUIREMENTS.exists():
        return ""
    return hashlib.sha256(REQUIREMENTS.read_bytes()).hexdigest()


def ensure_venv() -> Path:
    python = venv_python()
    if not python.exists():
        run([sys.executable, "-m", "venv", str(VENV)])
    return python


def ensure_dependencies(python: Path) -> None:
    current_hash = requirements_hash()
    previous_hash = STAMP.read_text(encoding="utf-8").strip() if STAMP.exists() else ""

    if current_hash == previous_hash:
        return

    run([str(python), "-m", "pip", "install", "-r", str(REQUIREMENTS)])
    STAMP.write_text(current_hash, encoding="utf-8")


def main() -> int:
    if getattr(sys, "frozen", False):
        from flashscore_parser.cli import main as cli_main

        return cli_main()

    python = ensure_venv()
    ensure_dependencies(python)
    return call([str(python), "-m", "flashscore_parser", *sys.argv[1:]])


if __name__ == "__main__":
    raise SystemExit(main())
