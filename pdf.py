"""
pdf.py

PDF and data logic for the THF Construction/FF planning grid.

Contains:
- Milestone, Task dataclasses
- Excel parsers (parse_excel, parse_manpower)
- PDF helpers (font registration, month colours, task shading)
- PDF generators:
    * generate_planning_grid
    * generate_planning_grid_with_manpower

Dependencies:
    pip install reportlab pandas openpyxl
"""

import os
import math
import datetime
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Dict, List, Tuple

import pandas as pd
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ============================================================
# Data structures
# ============================================================

@dataclass
class Milestone:
    """Represents a single milestone."""
    name: str
    date: date


@dataclass
class Task:
    """Represents a scheduled task for a contractor."""
    contractor: str
    name: str
    start_date: date
    duration_days: int


# ============================================================
# Excel parsing
# ============================================================

def parse_excel(filepath: str) -> tuple[List[Milestone], List[Task]]:
    """
    Parse the provided Excel file and extract milestones and tasks.

    Expected template structure (matching your shared file):
    - Milestones:
        Col "Unnamed: 1": name (header row "Milestones")
        Col "Unnamed: 2": date
    - Dynamic Motion:
        Cols "Unnamed: 4", "Unnamed: 5", "Unnamed: 6"  -> name, start, duration
    - MediaPro:
        Cols "Unnamed: 8", "Unnamed: 9", "Unnamed: 10" -> name, start, duration
    - Ocubo:
        Cols "Unnamed: 12","Unnamed: 13","Unnamed: 14" -> name, start, duration
    """

    df = pd.read_excel(filepath, sheet_name=0)

    milestones: List[Milestone] = []
    tasks: List[Task] = []

    def as_date(val):
        """Convert a cell value to a Python date if possible."""
        if pd.isna(val):
            return None
        try:
            return pd.to_datetime(val).date()
        except Exception:
            return None

    def add_task_from_row(row, name_col, start_col, dur_col, contractor_label: str):
        """Helper to read one contractor's task from a row."""
        name = row.get(name_col)
        start = row.get(start_col)
        dur = row.get(dur_col)

        if isinstance(name, str):
            name = name.strip()
        if not name or name == contractor_label:
            return

        start_d = as_date(start)
        if start_d is None or pd.isna(dur):
            return

        try:
            duration_days = int(dur)
        except Exception:
            return

        tasks.append(
            Task(
                contractor=contractor_label,
                name=name,
                start_date=start_d,
                duration_days=duration_days,
            )
        )

    for _, row in df.iterrows():
        # Milestones
        m_name = row.get("Unnamed: 1")
        m_day = row.get("Unnamed: 2")
        if isinstance(m_name, str):
            m_name = m_name.strip()
        if m_name and m_name != "Milestones":
            m_date = as_date(m_day)
            if m_date is not None:
                milestones.append(Milestone(name=m_name, date=m_date))

        # Tasks per contractor
        add_task_from_row(row, "Unnamed: 4", "Unnamed: 5", "Unnamed: 6", "Dynamic Motion")
        add_task_from_row(row, "Unnamed: 8", "Unnamed: 9", "Unnamed: 10", "MediaPro")
        add_task_from_row(row, "Unnamed: 12", "Unnamed: 13", "Unnamed: 14", "Ocubo")

    return milestones, tasks


def parse_manpower(filepath: str) -> tuple[Dict[date, float], Dict[str, Dict[date, float]], List[str]]:
    """
    Parse Dynamic Motion manpower from the second sheet.

    Layout (from screenshot):
      - Second sheet (sheet index 1).
      - One row contains the date headers: 11-Nov, 12-Nov, ...
      - One column to the left of the dates contains trade names
        (Foreman, Welding and Grinding, Gypsum Board, ...).
      - The grid under the date headers contains manpower counts.

    Returns:
        total_by_day: dict[date, float]                (sum over all trades)
        per_trade:   dict[trade_name, dict[date,float]]
        trade_order: list[str] in the order they appear in the sheet
    """
    try:
        # Read with header=None so no row is treated as column names
        df = pd.read_excel(filepath, sheet_name=1, header=None)
    except Exception:
        return {}, {}, []

    if df.empty:
        return {}, {}, []

    def to_date(val):
        if pd.isna(val):
            return None
        try:
            return pd.to_datetime(val).date()
        except Exception:
            return None

    # 1) Find the row that contains the date headers
    header_row_idx = None
    for i in range(len(df)):
        row_vals = df.iloc[i].tolist()
        date_like_count = sum(1 for v in row_vals if to_date(v) is not None)
        if date_like_count >= 3:
            header_row_idx = i
            break

    if header_row_idx is None:
        return {}, {}, []

    header_row = df.iloc[header_row_idx]

    # 2) Map column index -> date for the date columns
    col_date_map: Dict[int, date] = {}
    for col_idx, val in header_row.items():
        d = to_date(val)
        if d is not None:
            col_date_map[col_idx] = d

    if not col_date_map:
        return {}, {}, []

    # 3) Detect which column contains the trade names (left of first date col)
    first_date_col = min(col_date_map.keys())
    candidate_cols = [c for c in range(first_date_col) if c in df.columns]

    trade_col_idx = None
    best_score = -1

    for col in candidate_cols:
        col_series = df.iloc[header_row_idx + 1 :, col]
        score = 0
        for val in col_series:
            if isinstance(val, str) and val.strip():
                score += 1
        if score > best_score:
            best_score = score
            trade_col_idx = col

    if trade_col_idx is None:
        return {}, {}, []

    # 4) Build per-trade and total-by-day dictionaries
    total_by_day: Dict[date, float] = {}
    per_trade: Dict[str, Dict[date, float]] = {}
    trade_order: List[str] = []

    for row_idx in range(header_row_idx + 1, len(df)):
        row = df.iloc[row_idx]
        trade_name = row.iloc[trade_col_idx]

        if isinstance(trade_name, str):
            trade_name = trade_name.strip()
        else:
            continue

        if not trade_name:
            continue

        if trade_name not in per_trade:
            per_trade[trade_name] = {}
            trade_order.append(trade_name)

        for col_idx, d in col_date_map.items():
            val = row.iloc[col_idx]
            if pd.isna(val):
                continue
            try:
                f = float(val)
            except Exception:
                continue

            if f == 0:
                continue

            per_trade[trade_name][d] = per_trade[trade_name].get(d, 0.0) + f
            total_by_day[d] = total_by_day.get(d, 0.0) + f

    return total_by_day, per_trade, trade_order


# ============================================================
# PDF generation helpers
# ============================================================

def register_poppins_fonts(font_dir: str = ".") -> tuple[str, str]:
    """
    Register Poppins-Regular and Poppins-Bold with reportlab if present.

    Returns:
        (regular_font_name, bold_font_name)
    """
    regular_font_name = "Helvetica"
    bold_font_name = "Helvetica-Bold"

    regular_path = os.path.join(font_dir, "Poppins-Regular.ttf")
    bold_path = os.path.join(font_dir, "Poppins-Bold.ttf")

    try:
        if os.path.exists(regular_path):
            pdfmetrics.registerFont(TTFont("Poppins", regular_path))
            regular_font_name = "Poppins"
    except Exception:
        pass

    try:
        if os.path.exists(bold_path):
            pdfmetrics.registerFont(TTFont("Poppins-Bold", bold_path))
            bold_font_name = "Poppins-Bold"
    except Exception:
        pass

    return regular_font_name, bold_font_name


def get_month_colors() -> Dict[int, colors.Color]:
    """Stronger but still professional pastel-like colours for each month."""
    return {
        1: colors.HexColor("#CCE0FF"),
        2: colors.HexColor("#CFFFE0"),
        3: colors.HexColor("#FFE4C4"),
        4: colors.HexColor("#FFD6E8"),
        5: colors.HexColor("#E2D6FF"),
        6: colors.HexColor("#CFF7FF"),
        7: colors.HexColor("#E0F2B2"),
        8: colors.HexColor("#FFD1C7"),
        9: colors.HexColor("#D2D8FF"),
        10: colors.HexColor("#D4FFE2"),
        11: colors.HexColor("#FFD9B3"),
        12: colors.HexColor("#CDEBFF"),
    }


def make_shade_for_task(base_color: colors.Color, index: int, total: int) -> colors.Color:
    """
    Generate a clearly distinct shade for a task belonging to the same contractor.

    We vary both lightness and (slightly) the hue so that:
      - All colours still look like they belong to the same contractor family.
      - Adjacent tasks are much easier to distinguish.

    index: 0-based index of the task in that contractor's task list
    total: total number of tasks for that contractor
    """
    if total <= 1:
        return base_color

    import colorsys

    # Normalised position 0..1 across that contractor's tasks
    t = index / float(max(1, total - 1))

    # Base RGB from reportlab colour (0..1 range)
    r, g, b = base_color.red, base_color.green, base_color.blue

    # Convert to HLS so we can manipulate lightness and hue
    h, l, s = colorsys.rgb_to_hls(r, g, b)

    # Lightness: spread from "fairly dark" to "quite light"
    l_min, l_max = 0.35, 0.80
    l_new = l_min + (l_max - l_min) * t

    # Tiny hue shift either side of the base hue
    hue_shift_range = 0.06  # +/- 0.03
    h_new = (h + (t - 0.5) * hue_shift_range) % 1.0

    # Keep the same saturation so it still looks like the same colour family
    r2, g2, b2 = colorsys.hls_to_rgb(h_new, l_new, s)

    return colors.Color(r2, g2, b2)


# ============================================================
# Core PDF generator (with per-task shading)
# ============================================================

def generate_planning_grid(
    start_date: date,
    end_date: date,
    milestones: List[Milestone],
    tasks: List[Task],
    filename: str = "THF_Construction_FF_plan.pdf",
    cols: int = 7,
    margin_mm: float = 10.0,
    header_height_mm: float = 18.0,
) -> None:
    """
    Generate a planning grid PDF that includes milestones and contractor tasks.

    - Overlapping tasks for the SAME contractor are placed on different vertical lanes.
    - Lanes are grouped by contractor (Dynamic Motion, MediaPro, Ocubo, then any others).
    - Legend shows contractor colours.
    - Tasks for the same contractor are rendered in different shades of that contractor's colour.
    """
    if end_date < start_date:
        raise ValueError("end_date must be on or after start_date")

    num_days = (end_date - start_date).days + 1
    rows = math.ceil(num_days / cols)

    # --- Page setup ---
    page_width, page_height = landscape(A3)
    margin = margin_mm * mm
    header_height = header_height_mm * mm

    usable_width = page_width - 2 * margin
    usable_height = page_height - 2 * margin - header_height

    if usable_width <= 0 or usable_height <= 0:
        raise ValueError("Margins and header too large for page size")

    cell_width = usable_width / cols
    cell_height = usable_height / rows

    # --- Fonts for PDF ---
    regular_font_name, bold_font_name = register_poppins_fonts(".")
    title_font_name = bold_font_name

    c = canvas.Canvas(filename, pagesize=(page_width, page_height))
    c.setTitle("Ash's Works Planner")
    c.setAuthor("Ashley Pursglove")
    c.setSubject("Construction & FF Planning Grid")

    # --- Title / subtitle with version stamp ---
    title_text = "THF Construction/FF Plan"

    date_range_text = f"{start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')}"
    now = datetime.datetime.now()
    version_text = f" --- Version Generated at {now.strftime('%H:%M')} on {now.strftime('%d %b %Y')}"
    subtitle_text = date_range_text + version_text

    grid_top_y = margin + usable_height

    c.setFont(title_font_name, 18)
    title_width = c.stringWidth(title_text, title_font_name, 18)
    title_y = grid_top_y + header_height - 7 * mm
    c.drawString((page_width - title_width) / 2, title_y, title_text)

    c.setFont(regular_font_name, 12)
    subtitle_width = c.stringWidth(subtitle_text, regular_font_name, 12)
    subtitle_y = title_y - 5 * mm
    c.drawString((page_width - subtitle_width) / 2, subtitle_y, subtitle_text)

    # --- Grid origin ---
    grid_origin_x = margin
    grid_origin_y = margin

    month_colors = get_month_colors()
    weekend_color = colors.HexColor("#DDDDDD")

    total_cells = rows * cols
    current_date = start_date
    one_day = timedelta(days=1)

    # Map date -> cell bottom-left position
    date_positions: Dict[date, Tuple[float, float]] = {}

    # --- Backgrounds and date_positions ---
    for idx in range(total_cells):
        if current_date > end_date:
            break

        col = idx % cols
        row_from_top = idx // cols
        row = (rows - 1) - row_from_top

        x = grid_origin_x + col * cell_width
        y = grid_origin_y + row * cell_height

        weekday = current_date.weekday()  # Mon=0 .. Sun=6
        bg = month_colors.get(current_date.month, colors.white)
        if weekday in (4, 5):  # Fri, Sat
            bg = weekend_color

        c.setFillColor(bg)
        c.rect(x, y, cell_width, cell_height, stroke=0, fill=1)

        date_positions[current_date] = (x, y)
        current_date += one_day

    # --- Grid lines ---
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.3)
    c.rect(grid_origin_x, grid_origin_y, usable_width, usable_height)

    c.setLineWidth(0.7)
    for col in range(1, cols):
        x = grid_origin_x + col * cell_width
        c.line(x, grid_origin_y, x, grid_origin_y + usable_height)

    for row in range(1, rows):
        y = grid_origin_y + row * cell_height
        c.line(grid_origin_x, y, grid_origin_x + usable_width, y)

    # --------------------------------------------------------------
    # TASK LANE ASSIGNMENT (no overlaps per contractor)
    # AND PER-TASK COLOUR SHADES
    # --------------------------------------------------------------
    from collections import defaultdict as _dd

    contractor_task_indices = _dd(list)
    for idx, t in enumerate(tasks):
        contractor_task_indices[t.contractor].append(idx)

    slot_index_for_task: Dict[int, int] = {}
    max_slot_for_contractor: Dict[str, int] = {}

    for contractor, idxs in contractor_task_indices.items():
        idxs.sort(key=lambda i: tasks[i].start_date)
        active: List[tuple[date, int]] = []  # (end_date, slot)
        max_slot = -1

        for i in idxs:
            t = tasks[i]
            start_d = t.start_date
            end_d = t.start_date + timedelta(days=t.duration_days - 1)

            # drop finished tasks
            active = [a for a in active if a[0] >= start_d]
            used_slots = {slot for _, slot in active}

            slot = 0
            while slot in used_slots:
                slot += 1

            slot_index_for_task[i] = slot
            active.append((end_d, slot))
            max_slot = max(max_slot, slot)

        max_slot_for_contractor[contractor] = max_slot

    # contractor stacking order
    preferred_order = ["Dynamic Motion", "MediaPro", "Ocubo"]
    ordered_contractors: List[str] = []
    for c_name in preferred_order:
        if c_name in contractor_task_indices:
            ordered_contractors.append(c_name)
    for c_name in contractor_task_indices.keys():
        if c_name not in ordered_contractors:
            ordered_contractors.append(c_name)

    base_stack_for_contractor: Dict[str, int] = {}
    current_base = 0
    for contractor in ordered_contractors:
        base_stack_for_contractor[contractor] = current_base
        max_slot = max_slot_for_contractor.get(contractor, -1)
        if max_slot >= 0:
            current_base += max_slot + 1

    # visual settings
    contractor_colors = {
        "Dynamic Motion": colors.HexColor("#0077B6"),  # blue
        "MediaPro": colors.HexColor("#E63946"),        # red
        "Ocubo": colors.HexColor("#2A9D8F"),           # teal
    }

    # Build a per-task colour map with shaded variants per contractor
    task_color_map: Dict[int, colors.Color] = {}
    for contractor, idxs in contractor_task_indices.items():
        base_col = contractor_colors.get(contractor, colors.black)
        total = len(idxs)
        for pos, task_idx in enumerate(idxs):
            task_color_map[task_idx] = make_shade_for_task(base_col, pos, total)

    bar_height = 4 * mm
    bar_v_spacing = 1.5 * mm
    base_offset_from_bottom = 6 * mm
    bar_font_size = 9  # increased for readability

    # --------------------------------------------------------------
    # DRAW TASK BARS (label in each bar segment)
    # --------------------------------------------------------------
    for idx, task in enumerate(tasks):
        base_color = contractor_colors.get(task.contractor, colors.black)
        color = task_color_map.get(idx, base_color)

        task_slot = slot_index_for_task.get(idx, 0)
        contractor_base = base_stack_for_contractor.get(task.contractor, 0)
        stack_index = contractor_base + task_slot

        bar_label = task.name[:25]

        for day_offset in range(task.duration_days):
            d = task.start_date + timedelta(days=day_offset)
            if d < start_date or d > end_date:
                continue
            if d not in date_positions:
                continue

            cell_x, cell_y = date_positions[d]

            bar_y = (
                cell_y
                + base_offset_from_bottom
                + stack_index * (bar_height + bar_v_spacing)
            )
            bar_x = cell_x + 1.0 * mm
            bar_w = cell_width - 2.0 * mm

            c.setFillColor(color)
            c.setStrokeColor(color)
            c.rect(bar_x, bar_y, bar_w, bar_height, stroke=0, fill=1)

            c.setFont(regular_font_name, bar_font_size)
            c.setFillColor(colors.white)
            text_y = bar_y + (bar_height / 2.0) - (bar_font_size * 0.35)
            text_x = bar_x + 1.5 * mm
            c.drawString(text_x, text_y, bar_label)
            c.setFillColor(colors.black)

    # --------------------------------------------------------------
    # MILESTONES (dots + labels)
    # --------------------------------------------------------------
    milestones_by_date: Dict[date, List[Milestone]] = _dd(list)
    for ms in milestones:
        milestones_by_date[ms.date].append(ms)

    dot_radius = 2 * mm
    label_font_size = 8
    c.setFont(regular_font_name, label_font_size)
    vertical_spacing = dot_radius * 2 + 2 * mm

    for d, ms_list in milestones_by_date.items():
        if d < start_date or d > end_date:
            continue
        pos = date_positions.get(d)
        if not pos:
            continue

        cell_x, cell_y = pos
        start_cy = cell_y + cell_height - 6 * mm

        for i, ms in enumerate(ms_list):
            cy = start_cy - i * vertical_spacing
            cx = cell_x + cell_width - 4 * mm

            if cy - dot_radius < cell_y + 3 * mm:
                break

            c.setFillColor(colors.red)
            c.circle(cx, cy, dot_radius, stroke=1, fill=1)

            c.setFillColor(colors.black)
            label = ms.name[:30]
            text_y = cy - label_font_size * 0.35
            text_x = cx - (dot_radius + 2 * mm)
            c.drawRightString(text_x, text_y, label)

    # --------------------------------------------------------------
    # LEGEND FOR CONTRACTOR COLOURS
    # --------------------------------------------------------------
    legend_y = margin * 0.5       # centred in bottom margin
    legend_x = grid_origin_x
    legend_font_size = 8
    c.setFont(regular_font_name, legend_font_size)

    box_size = 5 * mm
    spacing_between_items = 8 * mm

    for contractor in ordered_contractors:
        color_box = contractor_colors.get(contractor, colors.black)

        c.setFillColor(color_box)
        c.setStrokeColor(colors.black)
        c.rect(legend_x, legend_y - box_size / 2, box_size, box_size, stroke=1, fill=1)

        c.setFillColor(colors.black)
        text_x = legend_x + box_size + 2 * mm
        text_y = legend_y - box_size / 4
        c.drawString(text_x, text_y, contractor)

        legend_x = (
            text_x
            + c.stringWidth(contractor, regular_font_name, legend_font_size)
            + spacing_between_items
        )

    # --------------------------------------------------------------
    # DATE LABELS (top-left of each cell)
    # --------------------------------------------------------------
    c.setFont(bold_font_name, 10)
    c.setFillColor(colors.black)

    current_date = start_date
    for idx in range(total_cells):
        if current_date > end_date:
            break
        cell_x, cell_y = date_positions[current_date]
        label = current_date.strftime("%a %d %b")
        c.drawString(cell_x + 3 * mm, cell_y + cell_height - 4 * mm, label)
        current_date += one_day

    # --------------------------------------------------------------
    # COPYRIGHT NOTICE
    # --------------------------------------------------------------
    c.setFont(regular_font_name, 8)
    copyright_text = (
        "© 2025 THF- Coded by Ashley Pursglove for THF. Source code and outputs are copyrighted. "
        "All rights reserved."
    )
    copyright_width = c.stringWidth(copyright_text, regular_font_name, 8)
    c.drawString((page_width - copyright_width) / 2, margin / 3, copyright_text)

    c.showPage()
    c.save()


def generate_planning_grid_with_manpower(
    start_date: date,
    end_date: date,
    milestones: List[Milestone],
    tasks: List[Task],
    manpower_by_day: Dict[date, float],
    manpower_by_trade: Dict[str, Dict[date, float]],
    trade_order: List[str],
    filename: str = "THF_Construction_FF_plan.pdf",
    cols: int = 7,
    margin_mm: float = 10.0,
    header_height_mm: float = 18.0,
) -> None:
    """
    Generate a 2-page PDF:
      - Page 1: planning grid (same style as generate_planning_grid, with shaded task bars).
      - Page 2: Dynamic Motion manpower summary & stacked histogram by trade.
    """
    # ---------- shared setup ----------
    page_width, page_height = landscape(A3)
    margin = margin_mm * mm
    header_height = header_height_mm * mm

    usable_width = page_width - 2 * margin
    usable_height = page_height - 2 * margin - header_height

    if usable_width <= 0 or usable_height <= 0:
        raise ValueError("Margins and header too large for page size")

    regular_font_name, bold_font_name = register_poppins_fonts(".")
    title_font_name = bold_font_name

    c = canvas.Canvas(filename, pagesize=(page_width, page_height))
    c.setTitle("Ash's Works Planner")
    c.setAuthor("Ashley Pursglove")
    c.setSubject("Construction & FF Planning Grid")

    # ==========================================================
    # PAGE 1: PLANNING GRID
    # ==========================================================
    num_days = (end_date - start_date).days + 1
    rows = math.ceil(num_days / cols)
    cell_width = usable_width / cols
    cell_height = usable_height / rows

    title_text = "THF Construction/FF Plan"

    date_range_text = f"{start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')}"
    now = datetime.datetime.now()
    version_text = f" --- Version Generated at {now.strftime('%H:%M')} on {now.strftime('%d %b %Y')}"
    subtitle_text = date_range_text + version_text

    grid_top_y = margin + usable_height

    c.setFont(title_font_name, 18)
    title_width = c.stringWidth(title_text, title_font_name, 18)
    title_y = grid_top_y + header_height - 7 * mm
    c.drawString((page_width - title_width) / 2, title_y, title_text)

    c.setFont(regular_font_name, 12)
    subtitle_width = c.stringWidth(subtitle_text, regular_font_name, 12)
    subtitle_y = title_y - 5 * mm
    c.drawString((page_width - subtitle_width) / 2, subtitle_y, subtitle_text)

    grid_origin_x = margin
    grid_origin_y = margin

    month_colors = get_month_colors()
    weekend_color = colors.HexColor("#DDDDDD")

    total_cells = rows * cols
    current_date = start_date
    one_day = timedelta(days=1)

    # Map date -> cell bottom-left position
    date_positions: Dict[date, Tuple[float, float]] = {}

    # Backgrounds + date_positions
    for idx in range(total_cells):
        if current_date > end_date:
            break

        col = idx % cols
        row_from_top = idx // cols
        row = (rows - 1) - row_from_top

        x = grid_origin_x + col * cell_width
        y = grid_origin_y + row * cell_height

        weekday = current_date.weekday()
        bg = month_colors.get(current_date.month, colors.white)
        if weekday in (4, 5):
            bg = weekend_color

        c.setFillColor(bg)
        c.rect(x, y, cell_width, cell_height, stroke=0, fill=1)

        date_positions[current_date] = (x, y)
        current_date += one_day

    # Grid lines
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.3)
    c.rect(grid_origin_x, grid_origin_y, usable_width, usable_height)

    c.setLineWidth(0.7)
    for col in range(1, cols):
        x = grid_origin_x + col * cell_width
        c.line(x, grid_origin_y, x, grid_origin_y + usable_height)

    for row in range(1, rows):
        y = grid_origin_y + row * cell_height
        c.line(grid_origin_x, y, grid_origin_x + usable_width, y)

    # ---------- task lane assignment ----------
    from collections import defaultdict as _dd

    contractor_task_indices = _dd(list)
    for idx, t in enumerate(tasks):
        contractor_task_indices[t.contractor].append(idx)

    slot_index_for_task: Dict[int, int] = {}
    max_slot_for_contractor: Dict[str, int] = {}

    for contractor, idxs in contractor_task_indices.items():
        idxs.sort(key=lambda i: tasks[i].start_date)
        active: List[tuple[date, int]] = []
        max_slot = -1
        for i in idxs:
            t = tasks[i]
            start_d = t.start_date
            end_d = t.start_date + timedelta(days=t.duration_days - 1)

            active = [a for a in active if a[0] >= start_d]
            used_slots = {slot for _, slot in active}

            slot = 0
            while slot in used_slots:
                slot += 1

            slot_index_for_task[i] = slot
            active.append((end_d, slot))
            max_slot = max(max_slot, slot)
        max_slot_for_contractor[contractor] = max_slot

    preferred_order = ["Dynamic Motion", "MediaPro", "Ocubo"]
    ordered_contractors: List[str] = []
    for c_name in preferred_order:
        if c_name in contractor_task_indices:
            ordered_contractors.append(c_name)
    for c_name in contractor_task_indices.keys():
        if c_name not in ordered_contractors:
            ordered_contractors.append(c_name)

    base_stack_for_contractor: Dict[str, int] = {}
    current_base_stack = 0
    for contractor in ordered_contractors:
        base_stack_for_contractor[contractor] = current_base_stack
        max_slot = max_slot_for_contractor.get(contractor, -1)
        if max_slot >= 0:
            current_base_stack += max_slot + 1

    contractor_colors = {
        "Dynamic Motion": colors.HexColor("#0077B6"),
        "MediaPro": colors.HexColor("#E63946"),
        "Ocubo": colors.HexColor("#2A9D8F"),
    }

    # Build per-task shaded colours
    task_color_map: Dict[int, colors.Color] = {}
    for contractor, idxs in contractor_task_indices.items():
        base_col = contractor_colors.get(contractor, colors.black)
        total = len(idxs)
        for pos, task_idx in enumerate(idxs):
            task_color_map[task_idx] = make_shade_for_task(base_col, pos, total)

    bar_height = 4 * mm
    bar_v_spacing = 1.5 * mm
    base_offset_from_bottom = 6 * mm
    bar_font_size = 9

    # Task bars (shaded per task)
    for idx, task in enumerate(tasks):
        base_color = contractor_colors.get(task.contractor, colors.black)
        color = task_color_map.get(idx, base_color)

        task_slot = slot_index_for_task.get(idx, 0)
        contractor_base = base_stack_for_contractor.get(task.contractor, 0)
        stack_index = contractor_base + task_slot
        bar_label = task.name[:25]

        for day_offset in range(task.duration_days):
            d = task.start_date + timedelta(days=day_offset)
            if d < start_date or d > end_date:
                continue
            if d not in date_positions:
                continue

            cell_x, cell_y = date_positions[d]
            bar_y = (
                cell_y
                + base_offset_from_bottom
                + stack_index * (bar_height + bar_v_spacing)
            )
            bar_x = cell_x + 1.0 * mm
            bar_w = cell_width - 2.0 * mm

            c.setFillColor(color)
            c.setStrokeColor(color)
            c.rect(bar_x, bar_y, bar_w, bar_height, stroke=0, fill=1)

            c.setFont(regular_font_name, bar_font_size)
            c.setFillColor(colors.white)
            text_y = bar_y + (bar_height / 2.0) - (bar_font_size * 0.35)
            text_x = bar_x + 1.5 * mm
            c.drawString(text_x, text_y, bar_label)
            c.setFillColor(colors.black)

    # Milestones
    milestones_by_date: Dict[date, List[Milestone]] = _dd(list)
    for ms in milestones:
        milestones_by_date[ms.date].append(ms)

    dot_radius = 2 * mm
    label_font_size = 8
    c.setFont(regular_font_name, label_font_size)
    vertical_spacing = dot_radius * 2 + 2 * mm

    for d, ms_list in milestones_by_date.items():
        if d < start_date or d > end_date:
            continue
        pos = date_positions.get(d)
        if not pos:
            continue

        cell_x, cell_y = pos
        start_cy = cell_y + cell_height - 6 * mm

        for i, ms in enumerate(ms_list):
            cy = start_cy - i * vertical_spacing
            cx = cell_x + cell_width - 4 * mm

            if cy - dot_radius < cell_y + 3 * mm:
                break

            c.setFillColor(colors.red)
            c.circle(cx, cy, dot_radius, stroke=1, fill=1)

            c.setFillColor(colors.black)
            label = ms.name[:30]
            text_y = cy - label_font_size * 0.35
            text_x = cx - (dot_radius + 2 * mm)
            c.drawRightString(text_x, text_y, label)

    # Legend (contractors)
    legend_y = margin * 0.5
    legend_x = grid_origin_x
    legend_font_size = 8
    c.setFont(regular_font_name, legend_font_size)

    box_size = 5 * mm
    spacing_between_items = 8 * mm

    for contractor in ordered_contractors:
        color_box = contractor_colors.get(contractor, colors.black)
        c.setFillColor(color_box)
        c.setStrokeColor(colors.black)
        c.rect(legend_x, legend_y - box_size / 2, box_size, box_size, stroke=1, fill=1)

        c.setFillColor(colors.black)
        text_x = legend_x + box_size + 2 * mm
        text_y = legend_y - box_size / 4
        c.drawString(text_x, text_y, contractor)

        legend_x = (
            text_x
            + c.stringWidth(contractor, regular_font_name, legend_font_size)
            + spacing_between_items
        )

    # Date labels for grid
    c.setFont(bold_font_name, 10)
    c.setFillColor(colors.black)
    current_date = start_date
    for idx in range(total_cells):
        if current_date > end_date:
            break
        cell_x, cell_y = date_positions[current_date]
        label = current_date.strftime("%a %d %b")
        c.drawString(cell_x + 3 * mm, cell_y + cell_height - 4 * mm, label)
        current_date += one_day

    # Copyright for page 1
    c.setFont(regular_font_name, 8)
    copyright_text = (
        "© 2025 THF- Coded by Ashley Pursglove for THF. Source code and outputs are copyrighted. "
        "All rights reserved."
    )
    copyright_width = c.stringWidth(copyright_text, regular_font_name, 8)
    c.drawString((page_width - copyright_width) / 2, margin / 3, copyright_text)

    c.showPage()

    # ==========================================================
    # PAGE 2: MANPOWER OVERVIEW + STACKED HISTOGRAM
    # ==========================================================
    title_text = "Dynamic Motion Manpower Overview"

    date_range_text = f"{start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')}"
    now = datetime.datetime.now()
    version_text = f" --- Version Generated at {now.strftime('%H:%M')} on {now.strftime('%d %b %Y')}"
    subtitle_text = date_range_text + version_text

    c.setFont(title_font_name, 18)
    title_width = c.stringWidth(title_text, title_font_name, 18)
    title_y = page_height - margin - 18 * mm
    c.drawString((page_width - title_width) / 2, title_y, title_text)

    c.setFont(regular_font_name, 12)
    subtitle_width = c.stringWidth(subtitle_text, regular_font_name, 12)
    subtitle_y = title_y - 6 * mm
    c.drawString((page_width - subtitle_width) / 2, subtitle_y, subtitle_text)

    # Build aligned lists for the selected date range
    dates_list = [start_date + i * one_day for i in range(num_days)]

    # Values per trade and per day
    values_per_trade: Dict[str, List[float]] = {}
    for trade in trade_order:
        per_day = manpower_by_trade.get(trade, {})
        values_per_trade[trade] = [float(per_day.get(d, 0.0)) for d in dates_list]

    # Total manpower per day (sum over trades)
    totals_per_day = [
        sum(values_per_trade[trade][i] for trade in trade_order)
        for i in range(num_days)
    ]
    max_val = max(totals_per_day) if totals_per_day else 0.0
    if max_val <= 0:
        max_val = 1.0

    total_man_days = sum(totals_per_day)
    working_days = sum(1 for v in totals_per_day if v > 0)
    avg_all_days = total_man_days / num_days if num_days > 0 else 0.0
    avg_working_days = total_man_days / working_days if working_days > 0 else 0.0
    peak = max(totals_per_day) if totals_per_day else 0.0
    peak_dates = [
        d for d, v in zip(dates_list, totals_per_day) if v == peak and v > 0
    ]

    # Metrics block (top-left)
    metrics_x = margin
    metrics_y = subtitle_y - 10 * mm
    c.setFont(regular_font_name, 10)
    line_h = 5 * mm

    c.drawString(metrics_x, metrics_y, f"Total man-days: {total_man_days:.1f}")
    c.drawString(metrics_x, metrics_y - line_h, f"Average manpower (all days): {avg_all_days:.2f}")
    c.drawString(metrics_x, metrics_y - 2 * line_h, f"Average manpower (working days): {avg_working_days:.2f}")
    c.drawString(metrics_x, metrics_y - 3 * line_h, f"Number of working days: {working_days}")

    if peak > 0:
        peak_dates_str = ", ".join(d.strftime("%d %b") for d in peak_dates)
        c.drawString(metrics_x, metrics_y - 4 * line_h, f"Peak manpower: {peak:.1f} on {peak_dates_str}")
    else:
        c.drawString(metrics_x, metrics_y - 4 * line_h, "Peak manpower: 0")

    # Trade legend (top-right, vertical list)
    legend_font_size = 9
    c.setFont(regular_font_name, legend_font_size)

    trade_palette = [
        "#FF7A18",  # orange
        "#00B894",  # green
        "#6C5CE7",  # purple
        "#0984E3",  # blue
        "#D63031",  # red
        "#E84393",  # pink
        "#2ECC71",  # light green
        "#F1C40F",  # yellow
    ]
    trade_colors: Dict[str, colors.Color] = {}
    for idx, trade in enumerate(trade_order):
        trade_colors[trade] = colors.HexColor(trade_palette[idx % len(trade_palette)])

    legend_x = page_width - margin - 50 * mm
    legend_y_top = metrics_y
    box_sz = 4 * mm
    legend_y = legend_y_top

    for trade in trade_order:
        col = trade_colors[trade]
        c.setFillColor(col)
        c.setStrokeColor(colors.black)
        c.rect(legend_x, legend_y - box_sz / 2, box_sz, box_sz, stroke=1, fill=1)

        c.setFillColor(colors.black)
        c.drawString(legend_x + box_sz + 2 * mm, legend_y - box_sz / 3, trade)
        legend_y -= 4.5 * mm

    # Histogram area
    chart_left = margin
    chart_right = page_width - margin
    chart_bottom = margin + 20 * mm
    chart_top = metrics_y - 8 * line_h
    chart_width = chart_right - chart_left
    chart_height = max(chart_top - chart_bottom, 40 * mm)

    # Axes
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.8)
    c.line(chart_left, chart_bottom, chart_right, chart_bottom)  # X axis
    c.line(chart_left, chart_bottom, chart_left, chart_bottom + chart_height)  # Y axis

    # Stacked bars per day
    bar_count = num_days
    if bar_count > 0:
        bar_spacing = chart_width / bar_count
        bar_actual_width = bar_spacing * 0.7

        for day_idx in range(bar_count):
            x = chart_left + day_idx * bar_spacing + (bar_spacing - bar_actual_width) / 2
            cumulative_height = 0.0

            for trade in trade_order:
                v = values_per_trade[trade][day_idx]
                if v <= 0:
                    continue

                segment_h = (v / max_val) * chart_height
                if segment_h <= 0:
                    continue

                segment_bottom = chart_bottom + cumulative_height

                c.setFillColor(trade_colors[trade])
                c.rect(
                    x,
                    segment_bottom,
                    bar_actual_width,
                    segment_h,
                    stroke=0,
                    fill=1,
                )

                if segment_h >= 3 * mm:
                    c.setFillColor(colors.black)
                    label_font_size = 7
                    c.setFont(regular_font_name, label_font_size)

                    if abs(v - int(v)) < 0.01:
                        label_text = f"{int(v)}"
                    else:
                        label_text = f"{v:.1f}"

                    label_x = x + bar_actual_width / 2
                    label_y = segment_bottom + segment_h / 2 - 2

                    text_w = c.stringWidth(label_text, regular_font_name, label_font_size)
                    c.drawString(label_x - text_w / 2, label_y, label_text)

                cumulative_height += segment_h

        # Y-axis labels: 0, max/2, max
        c.setFillColor(colors.black)
        c.setFont(regular_font_name, 8)
        for frac in (0.0, 0.5, 1.0):
            val = max_val * frac
            y = chart_bottom + chart_height * frac
            c.drawRightString(chart_left - 2 * mm, y - 2 * mm, f"{val:.0f}")

        # X-axis labels: label every day
        c.setFont(regular_font_name, 7)
        for day_idx, d in enumerate(dates_list):
            label = d.strftime("%d %b")
            lw = c.stringWidth(label, regular_font_name, 7)
            lx = chart_left + day_idx * bar_spacing + bar_spacing / 2 - lw / 2
            ly = chart_bottom - 5 * mm
            c.drawString(lx, ly, label)

    # Copyright on page 2
    c.setFont(regular_font_name, 8)
    copyright_text = (
        "© 2025 THF- Coded by Ashley Pursglove for THF. Source code and outputs are copyrighted. "
        "All rights reserved."
    )
    copyright_width = c.stringWidth(copyright_text, regular_font_name, 8)
    c.drawString((page_width - copyright_width) / 2, margin / 3, copyright_text)

    c.showPage()
    c.save()
