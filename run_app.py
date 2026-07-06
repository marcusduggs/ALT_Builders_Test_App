"""PyInstaller entry point for the PySide6 app.

ui/app.py does `from ui.main_window import MainWindow` -- an absolute
import of the top-level `ui` package, which only resolves when the repo
root is on sys.path. `python -m ui.app` arranges that automatically;
running `ui/app.py` directly (which is what PyInstaller needs -- it
can't target a `-m` module invocation) does not, since Python puts the
script's own directory (ui/) on sys.path, not its parent. This tiny
root-level wrapper is what PyInstaller actually builds against instead.
"""

from ui.app import main

if __name__ == "__main__":
    main()
