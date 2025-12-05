# """
# main.py

# Entry point for the THF Construction/FF Planner.

# Run:
#     python main.py
# """

# import sys
# from PyQt5 import QtWidgets

# from gui import apply_dark_theme, PlannerWindow


# def main() -> None:
#     """Create the QApplication, apply theme, and show the main window."""
#     app = QtWidgets.QApplication(sys.argv)
#     apply_dark_theme(app)

#     win = PlannerWindow()
#     win.show()

#     sys.exit(app.exec_())


# if __name__ == "__main__":
#     main()





"""
main.py

Entry point for Ash's Construction / FF Planner.

Creates the QApplication, applies the dark theme, and shows the main GUI.
"""

import sys
from PyQt5 import QtWidgets

from gui import PlannerWindow, apply_dark_theme  # type: ignore


def main() -> None:
    """
    Qt entry point.
    """
    app = QtWidgets.QApplication(sys.argv)
    apply_dark_theme(app)

    win = PlannerWindow()
    win.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
