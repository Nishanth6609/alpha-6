import os
from pathlib import Path
import sys


def set_window_icon(window):
    if os.name != "nt":
        return
    app_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    window.iconbitmap(default=str(app_dir / "ALPHA-6.ico"))
