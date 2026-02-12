import json
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_XLSX = REPO_ROOT / "source" / "project_dashboard.xlsx"
OUTPUT_JSON = REPO_ROOT / "data" / "risks.json"

# Accept either "Risk" or "Risks" tab name (handles your current workbook)
RISK_SHEET_CANDIDATES = ["Risks", "Risk"]


def find_sheet_name(xls: pd.ExcelFile, candidates: list[str]) -> str:
    for name in candidates:
        if name in xls.sheet_names:
            return name
    raise ValueError(f"None of these sheets exist: {candidates}. Found: {xls.sheet_names}")


def normalise_rating(x: str) -> str:
    """Normalise rating into Red/Amber/Green where possible."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    if not s:
        return ""
    s_low = s.lower()

    # Common mappings
    if s_low in {"red", "r"}:
        return "Red"
    if s_low in {"amber", "orange", "a", "medium", "med"}:
        return "Amber"
    if s_low in {"green", "g", "low"}:
        return "Green"
    if s_low in {"high"}:
        return "Red"

    # Fallback: title case
    return s[:1].upper() + s[1:].lower()


def to_iso_date(val) -> str:
    """Convert Excel date/datetime/cell value to YYYY-MM-DD string when possible."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    # pandas may give Timestamp
    if isinstance(val, pd.Timestamp):
        return val.date().isoformat()
    # datetime / date
    if hasattr(val, "date"):
        try:
            return val.date().isoformat()
        except Exception:
            pass
    # string
    s = str(val).strip()
    return s


def main():
    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"Missing source workbook: {SOURCE_XLSX}")

    xls = pd.ExcelFile(SOURCE_XLSX, engine="openpyxl")
    sheet = find_sheet_name(xls, RISK_SHEET_CANDIDATES)

    df = pd.read_excel(
        SOURCE_XLSX,
        sheet_name=sheet,
        engine="openpyxl",
        dtype=str,  # keep ids/text stable
    )

    # Expect these headers (you’ve already standardised these)
    expected = {"risk_id", "title", "status", "rating", "due_date", "owner_role", "last_updated"}
    missing = expected - set(df.columns)
    if missing:
        raise ValueError(f"Risks sheet is missing columns: {sorted(missing)}. Found: {list(df.columns)}")

    # Clean rows: drop completely empty ones
    df = df.dropna(how="all")

    items = []
    for _, row in df.iterrows():
        risk_id = (row.get("risk_id") or "").strip()
        title = (row.get("title") or "").strip()

        # Skip blank lines (common in Excel)
        if not risk_id and not title:
            continue

        item = {
            "id": risk_id,
            "description": title,  # dashboard uses "description" label; we map title -> description
            "status": (row.get("status") or "").strip(),
            "rating": normalise_rating(row.get("rating")),
            "ownerRole": (row.get("owner_role") or "").strip(),
            "nextActionDate": to_iso_date(row.get("due_date")),
            "lastUpdatedRow": to_iso_date(row.get("last_updated")),
        }
        items.append(item)

    now_awst = datetime.now(ZoneInfo("Australia/Perth")).strftime("%Y-%m-%d %H:%M AWST")
    out = {"lastUpdated": now_awst, "items": items}

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps(out, indent=2), encoding="utf-8")

    print(f"Wrote {OUTPUT_JSON} with {len(items)} items from sheet '{sheet}'")


if __name__ == "__main__":
    main()
