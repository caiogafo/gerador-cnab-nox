from __future__ import annotations

import subprocess
from pathlib import Path

FORBIDDEN_SUFFIXES = {".xls", ".xlsb", ".xlsm", ".pdf", ".mp4"}


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    result = subprocess.run(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard"],
        cwd=root,
        check=True,
        capture_output=True,
        text=True,
    )
    violations = []
    for item in result.stdout.splitlines():
        path = Path(item)
        if path.suffix.lower() in FORBIDDEN_SUFFIXES:
            violations.append(item)
        if path.suffix.lower() in {".xlsx", ".csv", ".txt"} and not str(path).startswith(
            "tests/fixtures/"
        ):
            violations.append(item)
    if violations:
        raise SystemExit("Arquivos potencialmente privados versionados: " + ", ".join(violations))
    print("privacy-check: OK")


if __name__ == "__main__":
    main()
