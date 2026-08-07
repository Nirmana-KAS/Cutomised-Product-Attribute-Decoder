# PROJECT STATUS — Customised Product Spec Decoder

## STATUS SUMMARY
| Field | Value |
|---|---|
| Application | Customised Product Spec Decoder |
| Current Version | v1.3 — 2026 |
| Total Edits | 4 |
| Last Edit Date | 2026-07-28 |
| Status | ✅ Active |

---

## EDIT-004 | 2026-07-28 | v1.3
**Title:** Major rewrite — individual products, 8-digit code, 3 screens, project-based duplicate resolution
**Requested By:** Developer (Boss correction)
**Description:** Complete redesign of processing logic based on corrected business requirements. Products are now individual (no grouping/combining). Name code is 8 digits with "F" prefix for finished products. Category changed from "Cut&Polished" to "Finish". Screen 4 removed (3 screens total). Internal Reference comes ONLY from Ref.No column (or Ref.No-Plan No. for 6 named projects). Mixed/MX/MIXED sizes are now skipped entirely. Duplicate rows resolved by selecting best row (most data filled) and collecting all LAB Certifications.
**Files Modified:** main.py (major rewrite), CLAUDE_CODE_INSTRUCTIONS.md (rewritten), PROJECT_STATUS.md
**Technical Changes:**
- Name code: "F" + stone_4digit + polish_1char + shape_2digit = 8 chars
- Category: "All/ Stones/ Finish / Stone Name"
- Removed Screen 4 — app now has 3 screens
- Individual products: no grouping, no comma-separated values, each row = one product
- Internal Reference: ONLY from Ref.No (or Ref.No-Plan No. for projects)
- 6 project names: #7-Nautilus, OVS Boucle, OVS Lunette, No. 3, Tache, Gala
- Duplicate resolution: group by product key, select best row, collect all LAB certs
- Internal Reference uniqueness enforced in export
- Skip: Mixed/MX/MIXED sizes, FS, None sizes, Metal variety
- Previously Decoded Products Library check still active
- Updated Polish Type Library (1,126 entries)
- Updated Category Name Library (Finish paths)
**Sections Affected:** All screens (Screen 2 rewritten, Screen 3 merged with old Screen 4, Screen 4 removed)
**Libraries Affected:** Category Name (Finish), Polish Type (updated), Previously Decoded Products (format changed)
**Dependencies:** None
**Related Apps:** None
**Data Impact:** Previously decoded products library MUST be cleared (format changed completely). app_data cleared during this edit — all 8 libraries need re-uploading; a v1.2 backup of the JSON files was taken before deletion.
**Rollback Risk:** High (major rewrite)
**Testing Result:** ✅ Passed — helper tests (format_size_value, size_to_range, build_ssku_with_size), PROJECT_NAMES = 6 entries, 26-column header layout (col 20 SSKU / 21 Plan No. / 22 Remarks / 23 LAB Cert / 24 Origin / 25-26 Upload Ref), build_internal_ref (project + Plan No., project without Plan No., non-project, blank Ref.No → blank), validate_row skips (blank/Metal variety, blank shape, Mixed/MX/MIXED, FS, blank size), resolve_duplicates (best row wins on score 3>2>1, LAB certs merged "CERT-A,CERT-B", upload rows merged "2,3,4", no-Ref.No rows never grouped), project vs non-project product keys, parse_size, match_category → Finish path, individual product rows (26 cols, 8-char F-prefixed name, single attribute values). End-to-end run on a synthetic 9-row invoice: 4 products / 4 skipped / 1 duplicate merged, refs "R1", "R9-No 1", "R9-No 2", "" — all unique; expanded 19 rows, compact 4 rows, both read back at 26 columns; auto-save added 4 entries and the re-run flagged 3 as already decoded. Source scan confirms zero Screen 4 / VariantsScreen / SummaryScreen references. GUI boots with exactly 3 screens and navigates 1→2→3→2→1 cleanly.
**Build Result:** ✅ Success — PyInstaller 6.16.0 / Python 3.14.0, dist\Customised_Product_Spec_Decoder.exe (30.2 MB)
**Success Rate:** 100% (all helper, integration and end-to-end assertions passed)
**Open Issues:** None — pending validation with the real updated libraries and a live customer invoice upload
**Claude Code Commands Used:** 4

## EDIT-003 | 2026-07-23 | v1.2
**Title:** Added LAB Certification and Origin display columns to all exports
**Requested By:** Developer
**Description:** Two new traceability columns added before the existing Upload Ref Sheet columns: "LAB Certification" (raw certificate numbers from upload) and "Origin" (resolved origin name from Origin Library). Origin Library updated to support both code matching (MM, SL, CH/AU) and full name matching (Pakistan, Sri Lanka, Madagascar/Tanzania). Both columns comma-separate values when multiple upload rows are grouped.
**Files Modified:** main.py, CLAUDE_CODE_INSTRUCTIONS.md, PROJECT_STATUS.md
**Technical Changes:**
- New field lab_certification stored per decoded product (raw value from upload)
- New field origin_display stored per decoded product (resolved via Origin Library lookup)
- Origin matching: case-insensitive exact match against "Origin Code" column, supports both codes and full names
- Origin display rules: unknown code, blank upload value, or blank library Origin Name → blank
- Screen 3 now 25 columns (cols 22-23 = LAB Cert + Origin, cols 24-25 = Upload Ref)
- Screen 4 now 26 columns (col 22 = SSKU, cols 23-24 = LAB Cert + Origin, cols 25-26 = Upload Ref)
- Grouped products: LAB Cert and Origin comma-separate all unique non-blank values
- Previously Decoded Products Library stores both new fields
- Search popup and Export include new fields
- Export styling applied to new columns (widths: LAB Certification 25, Origin 20)
- Skipped export verified — already contained raw LAB Certification and Origin columns
**Sections Affected:** Screen 2 (processing), Screen 3 (exports), Screen 4 (exports), Library 8
**Libraries Affected:** Origin Library (now accepts codes + full names), Library 8 (new stored fields)
**Dependencies:** None
**Related Apps:** None
**Data Impact:** Previously decoded products library should be cleared (new fields added)
**Rollback Risk:** Low
**Testing Result:** ✅ Passed — origin resolution (code match "SL"→Sri Lanka, full-name match "Pakistan"→Pakistan, "Madagascar/Tanzania"→Madagascar/Tanzania, unknown code→blank, blank library name→blank, blank upload→blank); Screen 3 = 25 cols with grouped comma-separated values ("GIA-111,GIA-222" / "Sri Lanka,Madagascar/Tanzania") on first row only; Screen 4 = 26 cols with per-variant values; Library 8 stores both fields; streamed 26-column export OK; GUI opens/cycles/closes cleanly
**Build Result:** ✅ Success — PyInstaller 6.16.0 / Python 3.14.0, dist\Customised_Product_Spec_Decoder.exe (30.2 MB)
**Success Rate:** 100% (all assertions passed)
**Open Issues:** None
**Claude Code Commands Used:** 4

## EDIT-002 | 2026-07-23 | v1.1
**Title:** Size formatting, SSKU range logic, library duplicates, upload traceability columns
**Requested By:** Developer
**Description:** Four key changes: (1) Library uploads no longer remove duplicates for Libraries 1-7 to preserve intentional spelling variations. (2) All size values formatted with exactly 2 decimal places throughout decoded output. (3) Complete rewrite of SSKU with Size column — L&W sizes converted to floor-based ranges, D&H products use Diameter→H range mapping, L&W with Height includes H range, upload ranges converted to standard company ranges. (4) Two new traceability columns added to all exports: "Upload Ref Sheet Name" (upload filename) and "Upload Ref Sheet Raw No." (row numbers from upload file).
**Files Modified:** main.py, CLAUDE_CODE_INSTRUCTIONS.md, PROJECT_STATUS.md
**Technical Changes:**
- Removed duplicate removal from library upload functions (Libraries 1-7 only)
- Added format_size_value() helper — formats all size numbers to 2 decimal places
- Added size_to_range() helper — converts any size value to floor-based company standard range (X.00-X.99)
- Added build_ssku_with_size() — generates SSKU with Size for all cases
- SSKU with Size L&W: code-L-range,W-range (+ ,H-range if height exists)
- SSKU with Size D&H: code-H-Drange (Diameter value → H range, actual Height ignored)
- SSKU with Size upload ranges: convert to standard range using floor of first number
- Size attribute values remain actual values with 2 decimal formatting
- New column "Upload Ref Sheet Name" — upload filename for traceability
- New column "Upload Ref Sheet Raw No." — row number(s) from upload file, comma-separated for grouped products
- Both new columns added to Screen 3 (23 cols), Screen 4 (24 cols), and Previously Decoded Products Library
- Skipped export already had filename and row number
**Sections Affected:** Screen 2 (processing + storage), Screen 3 (exports), Screen 4 (exports + SSKU), Library 8 (stored fields)
**Libraries Affected:** Libraries 1-7 (upload logic), Library 8 (new stored fields)
**Dependencies:** None
**Related Apps:** Size range format matches App 5 standard ranges
**Data Impact:** Previously decoded products library should be cleared before re-testing (format changed)
**Rollback Risk:** Low
**Testing Result:** ✅ Passed — format_size_value 9/9, size_to_range 14/14, build_ssku_with_size 6/6, plus integration checks: 23/24-column row shapes, traceability on first row only ("2,3" grouping), per-variant filename + row number, Library 8 stores both new fields, no-dedup library save, streamed 24-column export, GUI opens/cycles/closes cleanly
**Build Result:** ✅ Success — PyInstaller 6.16.0 / Python 3.14.0, dist\Customised_Product_Spec_Decoder.exe (30.2 MB)
**Success Rate:** 100% (all 29 helper assertions + all integration assertions passed)
**Open Issues:** None
**Claude Code Commands Used:** 6

## EDIT-001 | 2026-07-22 | v1.0
**Title:** Initial build — Complete application
**Requested By:** Developer
**Description:** Built the complete Customised Product Spec Decoder application with 4 screens, 8 libraries, upload processing, summary and expanded variant exports. Handles customer invoice-style files with complex shape/polish type matching logic.
**Files Modified:** main.py (created), CLAUDE_CODE_INSTRUCTIONS.md (created), requirements.txt (created), PROJECT_STATUS.md (created)
**Technical Changes:**
- Screen 1: 8 library cards with CRUD operations, special Previously Decoded Products library with universal search and export
- Screen 2: Upload processing with 12-step decoding logic including multi-level Polish Type matching
- Screen 3: Summary decoded export (Expanded, Compact, Skipped) with 21 columns
- Screen 4: Expanded variants export with SSKU with Size column (22 columns), diameter range conversion
- Auto-save decoded products to prevent duplicates
- Full ORAVA branding and styling
**Sections Affected:** All 4 screens
**Libraries Affected:** All 8 libraries
**Dependencies:** None (standalone app)
**Related Apps:** Similar export format to App 5 (Product Attributes Decoder)
**Data Impact:** New application, no existing data affected
**Rollback Risk:** Low (new app)
**Testing Result:** ✅ Passed — headless engine smoke test (stone/shape/polish matching, all 4 size cases, 7-digit code build, category, tags, grouping, 21/22-column row shapes, SSKU with Size incl. diameter range conversion, duplicate detection round-trip, streamed Excel export) + GUI test (all 4 screens open, navigate, and close cleanly)
**Build Result:** ✅ Success — PyInstaller 6.16.0 / Python 3.14.0, dist\Customised_Product_Spec_Decoder.exe (30.2 MB)
**Success Rate:** 100% (all smoke-test assertions passed; second-pass duplicate detection skipped 6/6 re-uploaded rows)
**Open Issues:** None — pending validation with real library files and customer invoice uploads
**Claude Code Commands Used:** 7
