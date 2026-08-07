# Customised Product Spec Decoder — CLAUDE_CODE_INSTRUCTIONS.md

**App Name:** Customised Product Spec Decoder
**Version:** v1.3 — 2026
**Folder:** customised_product_spec_decoder
**Exe:** Customised_Product_Spec_Decoder.exe
**Developer:** Shehan Nirmana
**Company:** ORAVA (Pvt) Ltd
**Footer:** "ORAVA (Pvt) Ltd | Developed by Shehan Nirmana | v1.3 — 2026"

## PURPOSE
Decode customised customer product orders into ERP (Odoo) product template format. Takes customer invoice files and produces individual product entries with standardised 8-digit codes.

## KEY RULES
1. ALL products are FINISHED products — Name code starts with "F" (8 digits total)
2. Every product is INDIVIDUAL — no grouping, no comma-separated values
3. Internal Reference comes ONLY from Ref.No column (never Order No., Remarks, etc.)
4. Mixed/MX/MIXED sizes are SKIPPED (cannot define BOM)
5. 6 named projects have shared Ref.No — use Plan No. to differentiate

## TECHNICAL STACK
- Python 3.14, tkinter + ttk GUI
- openpyxl for .xlsx read/write
- Pillow for logo (LANCZOS resize 1200x424 → 180x64)
- PyInstaller for .exe packaging
- JSON files in app_data folder (next to .exe, NOT in sys._MEIPASS)
- Libraries always loaded FRESH from disk
- White theme: BG=#FFFFFF, cards=#F0F2F5
- Logo centered, title centered below on every screen
- Window icon: icon.ico
- Write-only mode (openpyxl) for large exports with progress popup
- Preview tables capped at 2000 rows for performance
- Custom RoundButton canvas widget with hover effects
- Scrollable Treeview for data preview

## 3 SCREENS (not 4)

### Screen 1 — Library Manager (8 libraries)
8 library cards (scrollable). Each card: library name, description, entry count ("✔ X entries loaded" or "No data yet"). Logo centered at top, title below.

**Libraries 1-7:** Standard buttons — Upload, Edit (popup with search, add/edit/delete per row), Download, Delete All. Library uploads keep ALL rows — no duplicate removal (spelling variations are intentional).

**Library 8 (Previously Decoded Products):** Special — Export button, Universal Search button, Delete All button. Auto-populated when products are decoded. Stored as JSON.

#### Library 1: Stone Name (stone_library.json)
- 2 columns: "1-4 Digit" → "Stone name"
- ~115 entries
- Letters-only matching for lookups (strip digits from code, match on letters only)

#### Library 2: Shape - 2 Digit (shape_library.json)
- 3 columns: "Shape Code" → "Shape Name" → "Size Type"
- Size Type: "L & W" or "D & H"
- ~53 entries including TU (Tube), BE (Bead)

#### Library 3: Full Name Shape (fullname_shape_library.json)
- 4 columns: "Shape" → "Shape Code" → "Shape Name" → "Size Type"
- Maps long/variant shape names to 2-digit codes (~21 entries)
- Used when upload Shape column has more than 2 characters

#### Library 4: Polish Type Lookup (polish_type_library.json)
- 6 columns: "Shape", "Ref.No", "Plan No.", "Remarks" → "Polish Type Code", "Polishing Type Name"
- ~1,126 entries — massive lookup table
- Polish Type Code: single char (F, C, B, R, T, S, E)
- Contains duplicates intentionally for best matching
- THE ONLY AUTHORITY for determining polish type

#### Library 5: Origin (origin_library.json)
- 2 columns: "Origin Code" → "Origin Name"
- ~23+ entries
- Accepts BOTH short codes ("MM", "SL", "CH/AU") AND full names ("Pakistan", "Sri Lanka")

#### Library 6: Category Name (category_library.json)
- 2 columns: "Stone Name" → "Parent Category"
- Uses the "Finish" path for all stones: "All/ Stones/ Finish"
- Full category = Parent Category + " / " + Stone Name

#### Library 7: Colour (colour_library.json)
- 2 columns: "Stone Name" → "Colour"
- NOT used for decoding — Colour attribute always blank (manually filled later)

#### Library 8: Previously Decoded Products (decoded_products_library.json)
- Stores full decoded output for every exported product
- Export / Universal Search / Delete All
- Used to skip products already decoded in an earlier session

### Screen 2 — Upload & Process

**Upload file: 11 columns** (Item No., Order No., Ref.No, LAB Certification, Origin, Variety, Shape, Size, Height, Plan No., Remarks)

**Processing Steps:**

Step 1: Clean all values (strip whitespace, replace \xa0)
Step 2: Skip rows where Variety is blank/None or "Metal"
Step 3: Skip rows where Shape is blank/None
Step 4: Skip rows where Size is blank/None, "FS", "Mixed", "MX", "MIXED" (case-insensitive) — reason: "Mixed size - cannot define BOM" or "Invalid or missing size"
Step 5: Match Stone Name via Stone Name Library
Step 6: Match Shape (2-step: Shape Library then Full Name Shape Library)
Step 7: Match Polish Type (8-level priority matching via Polish Type Library)
Step 8: Match Origin via Origin Library
Step 9: Build 8-digit Name: "F" + stone_4digit + polish_1char + shape_2digit
Step 10: Parse Size (X format → L&W, single → D, range → D range). Format all to 2 decimals.
Step 11: Build Internal Reference (see below)
Step 12: Identify and resolve duplicates (see below)
Step 13: Check against Previously Decoded Products Library
Step 14: Build Tags, Category, SSKU with Size

**Polish Type match priority (highest to lowest — stop at first match):**
1. Shape + Ref.No + Plan No. + Remarks
2. Shape + Ref.No + Plan No.
3. Shape + Ref.No + Remarks
4. Shape + Ref.No
5. Shape + Plan No. + Remarks
6. Shape + Plan No.
7. Shape + Remarks
8. Shape only (if only ONE unique polish type exists for that shape)

If NO match at any level → SKIP, reason: "Polish type not found in library"

**Internal Reference Rules:**
- 6 PROJECT NAMES: #7-Nautilus, OVS Boucle, OVS Lunette, No. 3, Tache, Gala
- If Remarks matches a project name AND Ref.No has value AND Plan No. has value → Internal Reference = "Ref.No-Plan No." (e.g., "5711-104982-No 1", "VMMXE0Q6L3-P001")
- If Remarks matches a project name AND Ref.No has value BUT Plan No. is blank → Internal Reference = "Ref.No"
- If Remarks does NOT match project name AND Ref.No has value → Internal Reference = "Ref.No"
- If Ref.No is blank → Internal Reference = blank (regardless of Order No., Remarks, Plan No.)
- Internal Reference MUST be unique in exported file

**Duplicate Resolution:**
Products are grouped by their product key:
- Project products: (Ref.No + Plan No. + Variety + Shape + Size)
- Non-project products with Ref.No: (Ref.No + Variety + Shape + Size)
- Products without Ref.No: NO grouping, each row is individual

From each group, select the BEST row (most columns filled among Ref.No, Remarks, Plan No.):
Priority: Ref.No+Remarks+Plan No. (3) > Ref.No+Remarks (2) > Ref.No+Plan No. (2) > Ref.No (1)

LAB Certifications from ALL rows in the group are collected and comma-separated.

If two different products end up with same Internal Reference after processing → keep first, skip second.

### Screen 3 — Export (26 columns)

Each product = individual entry with single attribute values.

**Columns:**
1=Internal Reference, 2=Name (8-digit), 3=Product Category, 4=Tags, 5=Unit("Units"),
6=Product Attributes/Attribute, 7=Product Attributes/Values,
8=Sales("True"), 9=Purchase("True"), 10=Product Type("Goods"),
11=Invoicing Policy("Delivered quantities"), 12=Control Policy("on received quantity"),
13=Inventory tracking("True"), 14=Tracking("By  Lot"),
15=Sales Price(blank), 16=Sales Taxes(blank), 17=Cost(blank), 18=Purchase Taxes(blank),
19=Company("Orava (Pvt) Ltd"),
20=SSKU with Size, 21=Plan No., 22=Remarks,
23=LAB Certification, 24=Origin,
25=Upload Ref Sheet Name, 26=Upload Ref Sheet Raw No.

**Attributes per product (single values):**
1. Type = "F"
2. Size-L + Size-W (if L&W shape) OR Size-D (if D&H) + Size-H (if height available)
3. Treatment = "H"
4. Colour = blank

**Tags:** "Stone Name,Polish Type Name,Shape Name"

**Three export buttons:** Export Expanded (green), Export Compact (blue), Export Skipped (orange)

- Export Expanded: first row of each product has all 26 columns, subsequent attribute rows have only cols 6-7.
- Export Compact: one row per product with all columns.

**Export Filenames:**
- "Customised Product Spec - Decoded.xlsx"
- "Customised Product Spec - Compact.xlsx"
- "Customised Product Spec - Skipped.xlsx"

**Screen 3 layout:** Summary card ("Total Products: X | Skipped: Y"), preview table (cap 2000 rows), three export buttons, Back button (no Next — last screen).

**SSKU with Size:** Same range logic as v1.2 (floor-based ranges, D→H mapping for D&H shapes)

**Auto-save to Previously Decoded Products Library after export.**

## SIZE FORMATTING
- Attribute values: actual values with 2 decimal places (1.83→"1.83", 4→"4.00")
- SSKU with Size: floor-based ranges (1.83→"1.00-1.99")
- D&H shapes: Diameter→H range in SSKU (D=4.00 → "H-4.00-4.99")
- L&W with Height: all three get ranges in SSKU

## EXPORT STYLING
Headers: #4472C4 blue, white bold, Aptos Narrow. Alternating #D6E4F0/white rows. Freeze header. Write-only mode for large exports.

## GUI STYLING
- Header: blue #4472C4, white bold, Aptos Narrow
- Custom RoundButton canvas widget with hover effects
- Scrollable Treeview for data preview tables
- Cards with #F0F2F5 background
- White (#FFFFFF) main background

## PERSISTENCE & DEPLOYMENT
- JSON files in app_data folder next to .exe (NOT in sys._MEIPASS)
- Always load fresh from disk
- PyInstaller: pyinstaller --onefile --windowed --icon=icon.ico --add-data "logo.png;." --add-data "icon.ico;." --name "Customised_Product_Spec_Decoder" main.py
