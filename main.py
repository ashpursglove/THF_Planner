"""
main.py

Entry point for the THF Construction/FF Planner.

Run:
    python main.py
"""

import sys
from PyQt5 import QtWidgets

from gui import apply_dark_theme, PlannerWindow


def main() -> None:
    """Create the QApplication, apply theme, and show the main window."""
    app = QtWidgets.QApplication(sys.argv)
    apply_dark_theme(app)

    win = PlannerWindow()
    win.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
