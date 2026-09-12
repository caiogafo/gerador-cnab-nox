from __future__ import annotations

import os

from gerador_cnab_nox.gui import main

if __name__ == "__main__":
    if os.environ.get("CNAB_NOX_SMOKE") == "1":
        print("gui-entrypoint-import: OK")
    else:
        main()
