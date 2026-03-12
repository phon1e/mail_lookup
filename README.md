# Email Lookup Tool

A browser-based tool that matches Thai organization names from an input Excel file against a database Excel file, then outputs the result with emails appended.

No installation required. Runs entirely in the browser.

[![Live Demo](https://img.shields.io/badge/▶-Live%20Demo-blue)](https://phon1e.github.io/mail_lookup/)
---

## How It Works

1. Upload File 1 (input) containing short organization names
2. Upload File 2 (database) containing full organization names and their emails
3. Configure which columns to read from each file
4. Run matching — the tool outputs a new Excel file identical to File 1 with an email column added

---

## Matching Logic

Matching runs in two stages:

**Stage 1 — Substring match**
Checks if the input name is contained within the database name or vice versa. Fast and precise for cases where names differ only by appended district/province info.

Example:
- Input: `เทศบาลตำบลร่องคำ`
- Database: `เทศบาลตำบลร่องคำ อำเภอร่องคำ จังหวัดกาฬสินธุ์`
- Result: match found

**Stage 2 — Keyword scoring (fallback)**
If stage 1 finds nothing, the tool strips known Thai org-type prefixes from both sides and scores by shared keywords. The database entry with the highest keyword overlap is selected.

Example:
- Input: `เทศบาลตำบลยางม่วง`
- Database: `องค์การบริหารส่วนตำบลยางม่วง อำเภอท่ามะกา จังหวัดกาญจนบุรี`
- Stage 1: no substring match (different org-type prefix)
- Stage 2: strip prefixes, score on `ยางม่วง` — match found

Stripped prefixes include: เทศบาลนคร, เทศบาลเมือง, เทศบาลตำบล, องค์การบริหารส่วนตำบล, องค์การบริหารส่วนจังหวัด, อบต., อบจ., and others.

---

## Thai Character Normalization

Thai Excel files from different sources often encode the same characters differently. The tool normalizes all text before comparison:

- `ํา` (U+0E4D + U+0E32, two characters) is converted to `ำ` (U+0E33, single character)
- Whitespace is collapsed
- Unicode NFC normalization is applied

This prevents mismatches between visually identical names.

---

## Multiple Matches

When more than one database entry matches an input name, the tool picks the shortest database name (closest match). The match status column will note how many matches were found.

---

## Output

The output file is identical to File 1 with two columns appended:

| Column | Content |
|---|---|
| Email (Matched) | The matched email address, or blank if not found |
| Match Status | FOUND / FOUND (N matches) / NOT FOUND |

---

## Configuration

| Setting | Default | Description |
|---|---|---|
| File 1 org name column | C | Column containing short org names in the input file |
| File 2 name column | G | Column containing full org names in the database |
| File 2 email column | I | Column containing emails in the database |
| Skip header rows | 1 | Number of header rows to skip before data begins |

---

## File Format

Both files must be `.xlsx` or `.xls`. Column indices are zero-based internally but shown as letter labels (A, B, C...) in the UI.

---

## Limitations

- Keyword scoring may produce false positives when place names are common across regions (e.g. two districts both named บ้านใหม่). Review flagged rows manually.
- The tool does not connect to any external service. All processing happens locally in the browser. No data is uploaded anywhere.
