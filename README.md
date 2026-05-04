# Clinic Scoreboard -> JSON Converter

Converts `Scoreboard_Test.xlsx` into a structured, queryable `output.json`.



## How to run

```bash
pip install -r requirements.txt
python convert.py
```

Output is written to `output.json` in the same directory.  
Requires Python 3.10+.

---

## JSON shape

The file has three top-level keys:

```
{
  "meta":   { ... }   ← run metadata (source file, timestamp, counts)
  "schema": [ ... ]   ← flat list of all 125 metric descriptors
  "weeks":  [ ... ]   ← one record per data row (week)
}
```

### Why this shape?

The spreadsheet mixes two concerns: **column metadata** (what a metric is, who owns it, what focus area it belongs to) and **weekly values**. Flattening both into a single array of rows would bury the metadata inside every cell, making it noisy to query. Separating them means:

- A dashboard can load `schema` once to build a column picker.
- An LLM or script can iterate `weeks` and look up context from `schema` by key.
- Filtering by `focus`, `source`, or `role` is a single-pass scan of `schema`, not a parse of every row.

### Example week record

```json
{
  "week_ending": "2026-02-16",
  "metrics": {
    "Total Revenue - All Services": {
      "value": 40454.28,
      "section": null,
      "focus": "Financial",
      "source": "EMR",
      "role": "J",
      "target_note": null
    },
    "Answer Rate": {
      "value": 0.93,
      "section": "PHONE PERFORMANCE",
      "focus": "Answer",
      "source": "CallHero",
      "role": "J",
      "target_note": null
    }
  }
}
```

Duplicate metric names (e.g. `PVA (4 wk avg)` appears under PT, RMT, CHIRO, and Pelvic Health) are disambiguated with a `__N` suffix (`PVA (4 wk avg)__1`, `__2`, etc.). The `schema` array preserves the original `metric` string alongside the disambiguated `key`.

---

## Decisions about the messy bits

**Spacer columns** — The spreadsheet uses blank columns as visual dividers between sections. Columns with no header and no data in any row are silently dropped. 16 were removed.

**Merged cells** — Only one merged region exists (`AJ1:AP1`, the "PHONE PERFORMANCE" section label). The script resolves merges by reading the top-left cell value and propagating it to every column in the range, so each column in the merge gets the correct `section` value.

**Formulas** — `openpyxl` is opened with `data_only=False` (the default), which means formula cells are read as their formula string, not a cached value. Rather than silently storing `null` or a stale cached number, formula cells are stored as `"__formula__=<expr>"`. This makes it explicit to any consumer that the value requires Excel recalculation. If cached values are preferred, switch to `data_only=True` — the trade-off is that you lose the formula expression and may get stale data if the file was never saved after the last edit.

**Targets / row 7** — The target row is sparse and contains a mix of numeric targets, descriptive strings, and backtick placeholders. These are preserved as `target_note` strings on the column metadata rather than coerced into a typed `target` field — the formats are too inconsistent to normalise without domain knowledge.

**Hyperlinks in Source row** — Two source cells contain `=HYPERLINK(...)` formulas pointing to Google Sheets trackers. These surface as `__formula__` values in the `source` field of those columns. A future pass could extract the display text from the formula.

**Column A (date)** — Column A holds the `week_ending` date and doubles as a row label. It is promoted to the top-level `week_ending` field on each record and excluded from the `metrics` dict.

---

## What I'd do with another two hours

- **Recalculate formula values** : pipe the file through LibreOffice headless to get actual computed values instead of formula strings, then store both (`value` + `formula`).
- **Typed targets** : parse `target_note` into a structured `{ type: "percentage", value: 0.75 }` object where the pattern is unambiguous (e.g. "Target = 75%").
- **Validate output against schema** : add a JSON Schema or Pydantic model so the shape is machine-verifiable and breaking changes to the source file are caught at parse time.
- **Multi-clinic support** : the sheet appears to be a single clinic; if the real system has one sheet per clinic, a `clinic_id` field and a glob input pattern would handle the folder in one pass.
