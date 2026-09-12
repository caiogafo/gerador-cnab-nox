#!/usr/bin/env bash
set -euo pipefail

uv run --locked --no-sync ruff check .
CNAB_NOX_REQUIRE_GUI=1 uv run --locked --no-sync pytest --cov=gerador_cnab_nox --cov-report=term-missing
CNAB_NOX_SMOKE=1 uv run --locked --no-sync python scripts/launch_gui.py
uv run --locked --no-sync python scripts/check_repo_privacy.py
