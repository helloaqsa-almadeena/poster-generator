AQSA POSTER GENERATOR — WEBSITE
================================

HOW TO RUN:
-----------
1. Install Python 3.8+ if not already installed
2. Open terminal / command prompt in this folder
3. Run: pip install -r requirements.txt
4. Run: python app.py
5. Open browser: http://localhost:5000

FILES:
------
app.py               — Flask backend (all Word generation logic)
templates/index.html — Website frontend
requirements.txt     — Python dependencies

FEATURES:
---------
- A4 Portrait: Single product per page, large red price + د.إ symbol
- Shelf Labels: A5 Landscape, 6 labels per page in table format
- 2 Items: A5 Landscape, two products per page with unit
- Manual entry or Excel upload (.xlsx)
- Currency: د.إ or AED
- Downloads Word (.docx) file instantly

EXCEL FORMAT:
-------------
Columns required: Description | Price | Unit (only for 2-item type)
