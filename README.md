# Collecting Data As Per Soundex Name Percentage

## 1. Install Python

Install Python 3.10 or newer.

During Windows installation, enable:

    Add Python to PATH

## 2. Install required packages

Open Command Prompt in this folder and run:

    pip install -r requirements.txt

## 3. Run

Double-click `run_project.bat`, or run:

    python soundex_name_matching.py

## 4. GUI

The application asks for:

- Input Excel file
- Output Excel file
- Matching Merge Percentage (default 80)

The merge rule is:

    calculated matching percentage > selected threshold

So with the default value 80, only matches strictly above 80% create merge/group edges.

## 5. Output workbook

The generated workbook contains:

### Original Data
Original source rows/columns.

### Processed Data
Original columns plus:
- Unique ID
- Normalized Name
- Soundex Code
- Match Group

### Match Details
Pair-level comparison details, including:
- names
- Document IDs
- Unique IDs
- matching percentage
- component similarities
- Soundex codes
- threshold
- status
- reason

### Final Merged Data
One row for each matched group plus individual rows for records that did not belong to a multi-row group.

All source columns are handled dynamically. The code does not require a predefined list such as Passport, Bank Account, DL, etc.

## Important matching behavior

Soundex is supporting evidence only. The final percentage combines:
- order-aware full-name similarity
- token similarity
- first-name similarity
- last-name similarity
- middle-name similarity
- phonetic/Soundex similarity

Different nonblank values found in the same source column inside a matched group are not silently discarded. They are retained with ` | ` separators and the column is listed in `Conflict Columns`.

## Notes for large files

The application is designed for dynamic column counts, including 1,200+ columns. Very large row counts can still take substantial time because name matching requires comparisons. The GUI remains responsive while processing.
