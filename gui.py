# """
# gui.py

# PyQt5 GUI for the THF Construction/FF Planner.

# Contains:
# - apply_dark_theme(app): global dark-blue theme (matches cable tray app)
# - open_pdf_file(path): open generated PDF with system viewer
# - PlannerWindow: main window that wires Excel selection to PDF generation
# """

# import os
# import sys
# from datetime import date

# from PyQt5 import QtCore, QtWidgets
# from PyQt5.QtGui import QFontDatabase, QFont

# from pdf import (
#     parse_excel,
#     parse_manpower,
#     generate_planning_grid_with_manpower,
# )


# # ============================================================
# # Qt theming
# # ============================================================

# def apply_dark_theme(app: QtWidgets.QApplication):
#     """
#     Apply a dark blue style sheet with dayglow orange accent buttons and Poppins font
#     (if the TTF files are present next to this script or installed).

#     This matches the styling used in Ash's Cable Tray Calculator.
#     """
#     # Try to load Poppins into Qt's font DB from local files
#     script_dir = os.path.dirname(os.path.abspath(__file__))
#     for fname in ["Poppins-Regular.ttf", "Poppins-Bold.ttf"]:
#         fpath = os.path.join(script_dir, fname)
#         if os.path.exists(fpath):
#             QFontDatabase.addApplicationFont(fpath)

#     # Prefer Poppins if available, else fall back to system default
#     if "Poppins" in QFontDatabase().families():
#         app.setFont(QFont("Poppins", 10))
#     else:
#         app.setFont(QFont("Segoe UI", 10))

#     # Global dark-blue stylesheet
#     app.setStyleSheet("""
#     QWidget {
#         background-color: #0B1020;                  /* deep navy background */
#         color: #E5E7EB;                             /* soft near-white text */
#         font-family: "Poppins", "Segoe UI", sans-serif;
#         font-size: 10pt;
#     }

#     QMainWindow, QDialog {
#         background-color: #0B1020;
#     }

#     /* Panels / containers */
#     QFrame, QGroupBox {
#         background-color: #111827;                 /* slightly lighter panel bg */
#         border: 1px solid #1F2937;
#         border-radius: 6px;
#         margin-top: 6px;
#     }

#     QGroupBox::title {
#         subcontrol-origin: margin;
#         left: 8px;
#         padding: 0 4px;
#         color: #9CA3AF;
#         font-weight: 600;
#     }

#     QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QDateEdit {
#         background-color: #0F172A;                 /* input fields */
#         border: 1px solid #1F2937;
#         padding: 4px 6px;
#         border-radius: 4px;
#         color: #E5E7EB;
#         selection-background-color: #2563EB;
#         selection-color: #FFFFFF;
#     }

#     QLineEdit:disabled,
#     QSpinBox:disabled,
#     QDoubleSpinBox:disabled,
#     QComboBox:disabled,
#     QDateEdit:disabled {
#         color: #6B7280;
#         background-color: #020617;
#     }

#     QComboBox::drop-down {
#         border: none;
#         width: 18px;
#     }

#     QComboBox::down-arrow {
#         image: none;
#         width: 0;
#         height: 0;
#         border-left: 4px solid transparent;
#         border-right: 4px solid transparent;
#         border-top: 6px solid #9CA3AF;
#         margin-right: 4px;
#     }

#     QPushButton {
#         background-color: #FF7A18;                 /* dayglow orange */
#         color: #000000;
#         border: none;
#         border-radius: 4px;
#         padding: 6px 14px;
#         font-weight: 600;
#     }

#     QPushButton:hover {
#         background-color: #FF9A3D;                 /* lighter on hover */
#     }

#     QPushButton:pressed {
#         background-color: #E36800;                 /* darker when pressed */
#     }

#     QPushButton:disabled {
#         background-color: #4B5563;
#         color: #9CA3AF;
#     }

#     QScrollArea {
#         border: none;
#         background-color: #0B1020;
#     }

#     QScrollBar:vertical {
#         background: #020617;
#         width: 10px;
#         margin: 0px;
#     }
#     QScrollBar::handle:vertical {
#         background: #4B5563;
#         min-height: 20px;
#         border-radius: 4px;
#     }
#     QScrollBar::handle:vertical:hover {
#         background: #6B7280;
#     }
#     QScrollBar::add-line:vertical,
#     QScrollBar::sub-line:vertical {
#         height: 0;
#     }

#     QLabel#titleLabel {
#         font-size: 18pt;
#         font-weight: 600;
#         margin-bottom: 8px;
#         color: #F9FAFB;
#     }

#     QLabel#subtitleLabel {
#         font-size: 10pt;
#         color: #9CA3AF;
#         margin-bottom: 16px;
#     }

#     QLabel#statusLabel {
#         font-size: 9pt;
#         color: #9CA3AF;
#     }
#     """)


# # ============================================================
# # PDF opening helper
# # ============================================================

# def open_pdf_file(path: str) -> None:
#     """
#     Open the generated PDF with the default system viewer.

#     Works on:
#       - Windows  (os.startfile)
#       - macOS    (open)
#       - Linux    (xdg-open)
#     """
#     try:
#         if not os.path.isfile(path):
#             return

#         if os.name == "nt":  # Windows
#             os.startfile(path)  # type: ignore[attr-defined]
#         else:
#             import subprocess
#             if sys.platform == "darwin":  # macOS
#                 subprocess.Popen(["open", path])
#             else:  # Linux and others
#                 subprocess.Popen(["xdg-open", path])
#     except Exception:
#         # Fail silently – PDF is still generated even if auto-open fails
#         pass


# # ============================================================
# # Main planner window
# # ============================================================

# class PlannerWindow(QtWidgets.QWidget):
#     """
#     Dark-themed GUI to select an Excel file and generate the planning grid PDF.
#     Allows the user to pick a custom start and end date for the PDF.
#     """

#     def __init__(self, parent=None):
#         super().__init__(parent)
#         self.setWindowTitle("THF Construction/FF Planner")
#         self.resize(700, 500)
#         self.setMinimumSize(500, 300)

#         self.excel_path: str | None = None
#         self._build_ui()

#     def _build_ui(self) -> None:
#         main_layout = QtWidgets.QVBoxLayout(self)
#         main_layout.setContentsMargins(20, 20, 20, 20)
#         main_layout.setSpacing(12)

#         # Header
#         title = QtWidgets.QLabel("THF Construction/FF Planner")
#         title.setObjectName("titleLabel")

#         subtitle = QtWidgets.QLabel(
#             "Select the project Excel file, choose the date range, then generate a PDF."
#         )
#         subtitle.setObjectName("subtitleLabel")
#         subtitle.setWordWrap(True)

#         main_layout.addWidget(title)
#         main_layout.addWidget(subtitle)

#         # File selector row
#         file_row = QtWidgets.QHBoxLayout()
#         file_row.setSpacing(8)

#         lbl_file = QtWidgets.QLabel("Excel file:")
#         self.file_edit = QtWidgets.QLineEdit()
#         self.file_edit.setReadOnly(True)

#         browse_btn = QtWidgets.QPushButton("Browse Excel…")
#         browse_btn.clicked.connect(self.on_browse)

#         file_row.addWidget(lbl_file)
#         file_row.addWidget(self.file_edit)
#         file_row.addWidget(browse_btn)

#         main_layout.addLayout(file_row)

#         # Date range row
#         date_row = QtWidgets.QHBoxLayout()
#         date_row.setSpacing(8)

#         lbl_start = QtWidgets.QLabel("Start date:")
#         self.start_date_edit = QtWidgets.QDateEdit()
#         self.start_date_edit.setCalendarPopup(True)
#         self.start_date_edit.setDisplayFormat("dd MMM yyyy")
#         # Default: 4 Dec 2025
#         self.start_date_edit.setDate(QtCore.QDate(2025, 12, 4))

#         lbl_end = QtWidgets.QLabel("End date:")
#         self.end_date_edit = QtWidgets.QDateEdit()
#         self.end_date_edit.setCalendarPopup(True)
#         self.end_date_edit.setDisplayFormat("dd MMM yyyy")
#         # Default: 20 Dec 2025
#         self.end_date_edit.setDate(QtCore.QDate(2025, 12, 20))

#         date_row.addWidget(lbl_start)
#         date_row.addWidget(self.start_date_edit)
#         date_row.addSpacing(16)
#         date_row.addWidget(lbl_end)
#         date_row.addWidget(self.end_date_edit)
#         date_row.addStretch(1)

#         main_layout.addLayout(date_row)

#         # Generate button centred
#         self.generate_btn = QtWidgets.QPushButton("Generate PDF")
#         self.generate_btn.clicked.connect(self.on_generate)
#         self.generate_btn.setSizePolicy(
#             QtWidgets.QSizePolicy.Fixed,
#             QtWidgets.QSizePolicy.Fixed,
#         )

#         btn_row = QtWidgets.QHBoxLayout()
#         btn_row.addStretch(1)
#         btn_row.addWidget(self.generate_btn)
#         btn_row.addStretch(1)

#         main_layout.addLayout(btn_row)

#         # Status label
#         self.status_label = QtWidgets.QLabel("")
#         self.status_label.setObjectName("statusLabel")
#         self.status_label.setWordWrap(True)
#         main_layout.addWidget(self.status_label)

#         main_layout.addStretch(1)

#     # ------------- slots -------------

#     def on_browse(self) -> None:
#         """Open file dialog to select an Excel file."""
#         path, _ = QtWidgets.QFileDialog.getOpenFileName(
#             self,
#             "Select Excel File",
#             "",
#             "Excel Files (*.xlsx *.xls);;All Files (*)",
#         )
#         if path:
#             self.excel_path = path
#             self.file_edit.setText(path)
#             self.status_label.setText("")

#     def on_generate(self) -> None:
#         """Parse Excel and generate the PDF."""
#         if not self.excel_path:
#             self.status_label.setText("Please select an Excel file first.")
#             return

#         # Read date range from the UI
#         start_qdate = self.start_date_edit.date()
#         end_qdate = self.end_date_edit.date()
#         start = date(start_qdate.year(), start_qdate.month(), start_qdate.day())
#         end = date(end_qdate.year(), end_qdate.month(), end_qdate.day())

#         if end < start:
#             self.status_label.setText("End date must be on or after start date.")
#             return

#         try:
#             milestones, tasks = parse_excel(self.excel_path)
#             manpower_totals, manpower_by_trade, trade_order = parse_manpower(self.excel_path)

#             out_path = os.path.join(
#                 os.path.dirname(self.excel_path),
#                 "THF_Construction_FF_plan.pdf",
#             )

#             generate_planning_grid_with_manpower(
#                 start_date=start,
#                 end_date=end,
#                 milestones=milestones,
#                 tasks=tasks,
#                 manpower_by_day=manpower_totals,
#                 manpower_by_trade=manpower_by_trade,
#                 trade_order=trade_order,
#                 filename=out_path,
#             )

#             # Automatically open the generated PDF
#             open_pdf_file(out_path)

#             self.status_label.setText(
#                 f"PDF (grid + manpower) generated for "
#                 f"{start.strftime('%d %b %Y')} – {end.strftime('%d %b %Y')}:\n{out_path}"
#             )
#         except Exception as e:
#             self.status_label.setText(f"Error: {e}")





"""
gui.py

PyQt5 GUI for Ash's Construction / FF Planner.

Responsibilities:
- Let the user pick the Excel file.
- Let the user choose a start/end date.
- Let the user choose the output PDF filename.
- Call the PDF generation functions and open the resulting PDF.
"""

import os
from datetime import date

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QFontDatabase, QFont

from pdf import (
    parse_excel,
    parse_manpower,
    generate_planning_grid_with_manpower,
)


# ============================================================
# Theming / helpers
# ============================================================

def apply_dark_theme(app: QtWidgets.QApplication) -> None:
    """
    Apply a dark blue style sheet with dayglow orange accent buttons and Poppins font
    (if the TTF files are present next to this script or installed).
    This matches the styling used in Ash's Cable Tray Calculator.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Try to load Poppins into Qt's font DB from local files
    for fname in ["Poppins-Regular.ttf", "Poppins-Bold.ttf"]:
        fpath = os.path.join(script_dir, fname)
        if os.path.exists(fpath):
            QFontDatabase.addApplicationFont(fpath)

    # Prefer Poppins if available, else fall back to system default
    if "Poppins" in QFontDatabase().families():
        app.setFont(QFont("Poppins", 10))
    else:
        app.setFont(QFont("Segoe UI", 10))

    # Global dark-blue stylesheet (to match cable tray app)
    app.setStyleSheet("""
    QWidget {
        background-color: #0B1020;                  /* deep navy background */
        color: #E5E7EB;                             /* soft near-white text */
        font-family: "Poppins", "Segoe UI", sans-serif;
        font-size: 10pt;
    }

    QMainWindow, QDialog {
        background-color: #0B1020;
    }

    /* Panels / containers */
    QFrame, QGroupBox {
        background-color: #111827;                 /* slightly lighter panel bg */
        border: 1px solid #1F2937;
        border-radius: 6px;
        margin-top: 6px;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        left: 8px;
        padding: 0 4px;
        color: #9CA3AF;
        font-weight: 600;
    }

    QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QDateEdit {
        background-color: #0F172A;                 /* input fields */
        border: 1px solid #1F2937;
        padding: 4px 6px;
        border-radius: 4px;
        color: #E5E7EB;
        selection-background-color: #2563EB;
        selection-color: #FFFFFF;
    }

    QLineEdit:disabled,
    QSpinBox:disabled,
    QDoubleSpinBox:disabled,
    QComboBox:disabled,
    QDateEdit:disabled {
        color: #6B7280;
        background-color: #020617;
    }

    QComboBox::drop-down {
        border: none;
        width: 18px;
    }

    QComboBox::down-arrow {
        image: none;
        width: 0;
        height: 0;
        border-left: 4px solid transparent;
        border-right: 4px solid transparent;
        border-top: 6px solid #9CA3AF;
        margin-right: 4px;
    }

    QPushButton {
        background-color: #FF7A18;                 /* dayglow orange */
        color: #000000;
        border: none;
        border-radius: 4px;
        padding: 6px 14px;
        font-weight: 600;
    }

    QPushButton:hover {
        background-color: #FF9A3D;                 /* lighter on hover */
    }

    QPushButton:pressed {
        background-color: #E36800;                 /* darker when pressed */
    }

    QPushButton:disabled {
        background-color: #4B5563;
        color: #9CA3AF;
    }

    QScrollArea {
        border: none;
        background-color: #0B1020;
    }

    QScrollBar:vertical {
        background: #020617;
        width: 10px;
        margin: 0px;
    }
    QScrollBar::handle:vertical {
        background: #4B5563;
        min-height: 20px;
        border-radius: 4px;
    }
    QScrollBar::handle:vertical:hover {
        background: #6B7280;
    }
    QScrollBar::add-line:vertical,
    QScrollBar::sub-line:vertical {
        height: 0;
    }

    QLabel#titleLabel {
        font-size: 18pt;
        font-weight: 600;
        margin-bottom: 8px;
        color: #F9FAFB;
    }

    QLabel#subtitleLabel {
        font-size: 10pt;
        color: #9CA3AF;
        margin-bottom: 16px;
    }

    QLabel#statusLabel {
        font-size: 9pt;
        color: #9CA3AF;
    }
    """)


def open_pdf_file(path: str) -> None:
    """
    Open the generated PDF with the default system viewer.

    Works on:
      - Windows  (os.startfile)
      - macOS    (open)
      - Linux    (xdg-open)
    """
    try:
        if not os.path.isfile(path):
            return

        if os.name == "nt":  # Windows
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            import subprocess
            if sys.platform == "darwin":  # macOS
                subprocess.Popen(["open", path])
            else:  # Linux and others
                subprocess.Popen(["xdg-open", path])
    except Exception:
        # Fail silently – PDF is still generated even if auto-open fails
        pass


# ============================================================
# Main window
# ============================================================

class PlannerWindow(QtWidgets.QWidget):
    """
    Dark-themed GUI to select an Excel file and generate the planning grid PDF.
    Allows the user to pick a custom start and end date and a custom PDF filename.
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("THF Construction/FF Planner")
        self.resize(700, 500)
        self.setMinimumSize(500, 300)

        self.excel_path: str | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(12)

        # Header
        title = QtWidgets.QLabel("THF Construction/FF Planner")
        title.setObjectName("titleLabel")

        subtitle = QtWidgets.QLabel(
            "Select the project Excel file, choose the date range and PDF name, "
            "then generate a tidy, non-cursed PDF."
        )
        subtitle.setObjectName("subtitleLabel")
        subtitle.setWordWrap(True)

        main_layout.addWidget(title)
        main_layout.addWidget(subtitle)

        # File selector row
        file_row = QtWidgets.QHBoxLayout()
        file_row.setSpacing(8)

        lbl_file = QtWidgets.QLabel("Excel file:")
        self.file_edit = QtWidgets.QLineEdit()
        self.file_edit.setReadOnly(True)

        browse_btn = QtWidgets.QPushButton("Browse Excel…")
        browse_btn.clicked.connect(self.on_browse)

        file_row.addWidget(lbl_file)
        file_row.addWidget(self.file_edit)
        file_row.addWidget(browse_btn)

        main_layout.addLayout(file_row)

        # Date range row
        date_row = QtWidgets.QHBoxLayout()
        date_row.setSpacing(8)

        lbl_start = QtWidgets.QLabel("Start date:")
        self.start_date_edit = QtWidgets.QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("dd MMM yyyy")
        self.start_date_edit.setDate(QtCore.QDate(2025, 12, 4))  # default

        lbl_end = QtWidgets.QLabel("End date:")
        self.end_date_edit = QtWidgets.QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("dd MMM yyyy")
        self.end_date_edit.setDate(QtCore.QDate(2025, 12, 20))  # default

        date_row.addWidget(lbl_start)
        date_row.addWidget(self.start_date_edit)
        date_row.addSpacing(16)
        date_row.addWidget(lbl_end)
        date_row.addWidget(self.end_date_edit)
        date_row.addStretch(1)

        main_layout.addLayout(date_row)

        # PDF name row
        name_row = QtWidgets.QHBoxLayout()
        name_row.setSpacing(8)

        lbl_name = QtWidgets.QLabel("PDF file name:")
        self.name_edit = QtWidgets.QLineEdit()
        self.name_edit.setPlaceholderText("THF_Construction_FF_plan")
        self.name_edit.setText("THF_Construction_FF_plan")

        name_row.addWidget(lbl_name)
        name_row.addWidget(self.name_edit)

        main_layout.addLayout(name_row)

        # Generate button centred
        self.generate_btn = QtWidgets.QPushButton("Generate PDF")
        self.generate_btn.clicked.connect(self.on_generate)
        self.generate_btn.setSizePolicy(
            QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed
        )

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(self.generate_btn)
        btn_row.addStretch(1)

        main_layout.addLayout(btn_row)

        # Status label
        self.status_label = QtWidgets.QLabel("")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        main_layout.addWidget(self.status_label)

        main_layout.addStretch(1)

    # --------------------------------------------------------
    # Slots
    # --------------------------------------------------------

    def on_browse(self) -> None:
        """
        Open file dialog to select an Excel file.
        """
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)",
        )
        if path:
            self.excel_path = path
            self.file_edit.setText(path)
            self.status_label.setText("")

    def on_generate(self) -> None:
        """
        Parse Excel and generate the PDF.
        """
        if not self.excel_path:
            self.status_label.setText("Please select an Excel file first.")
            return

        # Read date range from the UI
        start_qdate = self.start_date_edit.date()
        end_qdate = self.end_date_edit.date()

        start = date(start_qdate.year(), start_qdate.month(), start_qdate.day())
        end = date(end_qdate.year(), end_qdate.month(), end_qdate.day())

        if end < start:
            self.status_label.setText("End date must be on or after start date.")
            return

        # Build output filename from user input
        raw_name = self.name_edit.text().strip()
        if not raw_name:
            raw_name = "THF_Construction_FF_plan"

        # Ensure .pdf extension
        if not raw_name.lower().endswith(".pdf"):
            raw_name = raw_name + ".pdf"

        # This is what we'll show in the subtitle after the version stamp
        display_name = os.path.splitext(os.path.basename(raw_name))[0]

        out_path = os.path.join(
            os.path.dirname(self.excel_path),
            raw_name,
        )

        try:
            milestones, tasks = parse_excel(self.excel_path)
            manpower_totals, manpower_by_trade, trade_order = parse_manpower(self.excel_path)

            generate_planning_grid_with_manpower(
                start_date=start,
                end_date=end,
                milestones=milestones,
                tasks=tasks,
                manpower_by_day=manpower_totals,
                manpower_by_trade=manpower_by_trade,
                trade_order=trade_order,
                filename=out_path,
                version_label=display_name,   # <- NEW: pass the name into the PDF
            )

            # Automatically open the generated PDF
            open_pdf_file(out_path)

            self.status_label.setText(
                f"PDF (grid + manpower) generated for "
                f"{start.strftime('%d %b %Y')} – {end.strftime('%d %b %Y')}:\n{out_path}"
            )
        except Exception as e:
            self.status_label.setText(f"Error: {e}")
