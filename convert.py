"""
convert.py  Clinic Scoreboard XLSX -> JSON converter
Reads "Scoreboard Test.xlsx" and writes output.json
"""

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

# config
INPUT_FILE = "Scoreboard Test.xlsx"
OUTPUT_FILE = "output.json"

# Row indices (1-based)
ROW_SECTION   = 1   # sparse section-header row (for example "PHONE PERFORMANCE")
ROW_METRIC    = 2   # main metric name
ROW_FOCUS     = 3   # focus category (Financial, Marketing, …)
ROW_SOURCE    = 4   # data source (EMR, CallHero, …)
ROW_ROLE      = 5   # responsible person / role
ROW_TARGETS   = 7   # inline target notes (sparse)
DATA_ROWS_START = 8 # first actual data row


def clean_header(value: str) -> str:
    """Strip newlines and extra whitespace from a cell header."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


def resolve_formula(cell):
    """
    Return a string tag for formula cells so consumers know the value
    was not pre-computed. Raw data values are returned as-is.
    """
    v = cell.value
    if v is None:
        return None
    if isinstance(v, str) and v.startswith("="):
        return f"__formula__{v}"
    return v


def coerce_value(v):
    """JSON-safe coercion: datetime → ISO string, everything else as-is."""
    if isinstance(v, datetime):
        return v.date().isoformat()
    return v


def build_column_metadata(ws):
    """
    Walk every column and return a dict keyed by column index with:
      section, metric, focus, source, role, target_note
    Skips columns that have no metric name AND no data in any data row.
    """
    max_col = ws.max_column
    data_rows = list(range(DATA_ROWS_START, ws.max_row + 1))

    # Resolve merged-cell section labels so every col inside the merge gets
    # the same section string.
    section_by_col = {}
    for merged_range in ws.merged_cells.ranges:
        top_val = ws.cell(merged_range.min_row, merged_range.min_col).value
        if merged_range.min_row == ROW_SECTION and top_val:
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                section_by_col[col] = clean_header(top_val)

    columns = {}
    for col in range(1, max_col + 1):
        metric = clean_header(ws.cell(ROW_METRIC, col).value)
        focus  = clean_header(ws.cell(ROW_FOCUS,  col).value)
        source = clean_header(ws.cell(ROW_SOURCE, col).value)
        role   = clean_header(ws.cell(ROW_ROLE,   col).value)
        target = clean_header(ws.cell(ROW_TARGETS, col).value)
        section = section_by_col.get(col, "")

        # Skip pure spacer columns
        has_data = any(ws.cell(r, col).value is not None for r in data_rows)
        if not metric and not has_data:
            continue

        # Skip col A (date/label column, not a metric)
        if col == 1:
            continue

        columns[col] = {
            "section":     section  or None,
            "metric":      metric   or None,
            "focus":       focus    or None,
            "source":      source   or None,
            "role":        role     or None,
            "target_note": target   or None,
        }

    return columns


def build_records(ws, columns):
    """
    Build one record per data row. Each record has a `week_ending` date and
    a `metrics` dict keyed by metric name (with collision suffix for dupes).
    """
    records = []

    # Build unique metric keys (some names repeat, for example "PVA (4 wk avg)")
    metric_key_counts = {}
    col_to_key = {}
    for col, meta in columns.items():
        raw = meta["metric"] or f"col_{col}"
        metric_key_counts[raw] = metric_key_counts.get(raw, 0) + 1

    seen = {}
    for col, meta in columns.items():
        raw = meta["metric"] or f"col_{col}"
        if metric_key_counts[raw] > 1:
            seen[raw] = seen.get(raw, 0) + 1
            key = f"{raw}__{seen[raw]}"
        else:
            key = raw
        col_to_key[col] = key

    for row in range(DATA_ROWS_START, ws.max_row + 1):
        date_cell = ws.cell(row, 1).value
        if date_cell is None:
            continue  # skip empty trailing rows

        record = {
            "week_ending": coerce_value(date_cell),
            "metrics": {}
        }

        for col, meta in columns.items():
            raw_val = resolve_formula(ws.cell(row, col))
            value   = coerce_value(raw_val)
            key     = col_to_key[col]

            record["metrics"][key] = {
                "value":       value,
                "section":     meta["section"],
                "focus":       meta["focus"],
                "source":      meta["source"],
                "role":        meta["role"],
                "target_note": meta["target_note"],
            }

        records.append(record)

    return records


def build_schema(columns, col_to_key=None):
    """
    Return a flat list of column descriptors — useful for documentation /
    dashboard column pickers without having to inspect a data record.
    """
    schema = []
    # Need col_to_key
    if col_to_key is None:
        metric_key_counts = {}
        for col, meta in columns.items():
            raw = meta["metric"] or f"col_{col}"
            metric_key_counts[raw] = metric_key_counts.get(raw, 0) + 1
        seen = {}
        col_to_key = {}
        for col, meta in columns.items():
            raw = meta["metric"] or f"col_{col}"
            if metric_key_counts[raw] > 1:
                seen[raw] = seen.get(raw, 0) + 1
                col_to_key[col] = f"{raw}__{seen[raw]}"
            else:
                col_to_key[col] = raw

    for col, meta in columns.items():
        schema.append({
            "key":         col_to_key[col],
            "metric":      meta["metric"],
            "section":     meta["section"],
            "focus":       meta["focus"],
            "source":      meta["source"],
            "role":        meta["role"],
            "target_note": meta["target_note"],
        })
    return schema


def main():
    src = Path(INPUT_FILE)
    if not src.exists():
        print(f"ERROR: {INPUT_FILE} not found.", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(src, data_only=False)
    ws = wb.active

    columns = build_column_metadata(ws)

    # Rebuild col_to_key consistently for both records and schema
    metric_key_counts = {}
    for col, meta in columns.items():
        raw = meta["metric"] or f"col_{col}"
        metric_key_counts[raw] = metric_key_counts.get(raw, 0) + 1
    seen = {}
    col_to_key = {}
    for col, meta in columns.items():
        raw = meta["metric"] or f"col_{col}"
        if metric_key_counts[raw] > 1:
            seen[raw] = seen.get(raw, 0) + 1
            col_to_key[col] = f"{raw}__{seen[raw]}"
        else:
            col_to_key[col] = raw

    records = build_records(ws, columns)
    schema  = build_schema(columns, col_to_key)

    output = {
        "meta": {
            "source_file":   src.name,
            "sheet":         ws.title,
            "generated_at":  datetime.now(timezone.utc).isoformat(),
            "total_weeks":   len(records),
            "total_metrics": len(schema),
        },
        "schema": schema,
        "weeks":  records,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"✓  Wrote {OUTPUT_FILE}  ({len(records)} week(s), {len(schema)} metrics)")


if __name__ == "__main__":
    main()