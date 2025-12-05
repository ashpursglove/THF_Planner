# THF_Planner
# Ash’s Construction / FF Planner
## A scheduling tool nobody asked for, but everybody definitely needed.

# What Is This?

## This glorious contraption is Ash’s Construction & FF Planner, a desktop tool built to answer one simple question:

#### “How do I turn a cursed Excel schedule into a beautiful, professional PDF… without crying?”

It takes the same spaghetti-logic planning spreadsheets everyone in construction pretends to understand, and turns them into:

A clean, colour-coded daily grid

Automatically stacked multi-lane task bars (no overlaps, no chaos)

Milestones with red marker dots

A manpower histogram (because numbers make you look serious)

A two-page PDF that looks like it was prepared by someone who definitely charges consultancy rates

And all of this happens with one button press, powered by Python, caffeine, and sheer spite.

# Features Nobody Should Have To Implement, Yet Here We Are

= Dark blue GUI theme matching the Cable Tray Calculator (unified brand identity = professional™)

= Date picker pre-filled to 4–20 Dec 2025 (because… those are your current problems)

= Automatic PDF generation in two pages:

= Page 1 → construction grid with month colours, Fri/Sat greys, contractor task blocks, and milestone dots

= Page 2 → manpower overview + stacked histogram by trade

= Shaded variations of each contractor colour so tasks don’t blend into one giant blue rectangle (“Dynamic Motion Blue Hell”)

= PDF auto-opens after generation

= Zero Excel formatting assumptions except the known template (your usual “Unnamed: X” art piece)

# How It Works (in Human Terms)

You point the app at your Excel file.

It rummages through the sheet, whispering “why… why was this built like this…?”

It reconstructs the timeline, sorts contractors, assigns slots, renders the grid, and decides where to put all the coloured rectangles.

It draws a second page showing how many humans need to be on site
(and reveals just how overworked the welding team actually is).

A shiny PDF appears on your screen like a gift from the scheduling gods.

Installation

You need Python 3.10+ and the following dependencies:

pip install pyqt5 reportlab pandas openpyxl


Optional but recommended:

Poppins font files (drop Poppins-Regular.ttf and Poppins-Bold.ttf in the same folder)

Running the Thing
python main.py


If the program doesn’t start, you probably:

forgot to install PyQt5

misspelled the file path

sacrificed the wrong type of incense to the Excel gods

Project Structure

Now beautifully modular:

project_root/
│
├── main.py           # Entry point — fires up the GUI
├── gui.py            # All PyQt5 UI logic
├── pdf.py            # Milestones, lanes, stacked bars, manpower page
├── Poppins-Regular.ttf
├── Poppins-Bold.ttf
└── README.md         # You're reading this masterpiece

The Excel Template (a necessary evil)

Sheet 1 should contain:

Milestones in columns 1–2

Dynamic Motion tasks in cols 4–6

MediaPro tasks in cols 8–10

Ocubo tasks in cols 12–14

Sheet 2 should contain:

A row of dates

A column of trades

A grid of numbers vaguely representing the number of humans

If your Excel file deviates from this, the app will respond with:

“Error: what fresh hell is this?”

Output Example

Page 1:
✔ Month colours
✔ Weekend shading
✔ Task bars (shaded per task)
✔ Milestone dots
✔ Contractor legend
✔ Date labels
✔ Copyright
✔ Overall smugness

Page 2:
✔ Total man-days
✔ Average manpower
✔ Peak manpower & dates
✔ Stacked histogram
✔ Trade legend
✔ Slight panic when you realise how much manpower is needed

Why Does This Exist?

Because:

Excel is a war crime.

Time is finite.

PDFs are pretty.

You needed something faster to present than a panicked screenshot during a client call.

License
© 2025 THF — Coded by Ashley Pursglove.
All rights reserved.
Unauthorized duplication will result in a strongly worded PDF.

