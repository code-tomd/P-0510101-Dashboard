import json
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_XLSX = REPO_ROOT / "source" / "project_dashboard.xlsx"
OUTPUT_JSON = REPO_ROOT / "data" / "risks.json"

# Accept either "Risk" or "Risks" tab name
RISK_SHEET_CANDIDATES = ["Risks", "Risk"]


def find_sheet_name(xls: pd.ExcelFile, candidates: list[str]) -> str:
    for name in candidates:
        if name in xls.sheet_names:
            return name
    raise ValueError(
        f"None of these sheets exist: {candidates}. Found: {xls.sheet_names}"
    )


def as_text(val) -> str:
    """Safe string conversion for Excel cells (handles NaN/None/numbers)."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val)


def normalise_rating(x) -> str:
    """Normalise rating into Red/Amber/Green where possible."""
    s = as_text(x).strip()
    if not s:
        return ""
    s_low = s.lower()

    # Common mappings
    if s_low in {"red", "r", "high"}:
        return "Red"
    if s_low in {"amber", "orange", "a", "medium", "med"}:
        return "Amber"
    if s_low in {"green", "g", "low"}:
        return "Green"

    # Fallback: Title Case
    return s[:1].upper() + s[1:].lower()


def to_iso_date(val) -> str:
    """Convert Excel date/datetime/cell value to YYYY-MM-DD string when possible."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass

    # pandas may provide Timestamp
    if isinstance(val, pd.Timestamp):
        return val.date().isoformat()

    # datetime/date-like
    if hasattr(val, "date"):
        try:
            return val.date().isoformat()
        except Exception:
            pass

    # string fallback
    return as_text(val).strip()


def main():
    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"Missing source workbook: {SOURCE_XLSX}")

    xls = pd.ExcelFile(SOURCE_XLSX, engine="openpyxl")
    sheet = find_sheet_name(xls, RISK_SHEET_CANDIDATES)

    df = pd.read_excel(
        SOURCE_XLSX,
        sheet_name=sheet,
        engine="openpyxl",
        dtype=object,  # keep dates as dates; keep ids stable
    )

    expected = {
        "risk_id",
        "title",
        "status",
        "rating",
        "due_date",
        "owner_role",
        "last_updated",
    }
    missing = expected - set(df.columns)
    if missing:
        raise ValueError(
            f"Risks sheet is missing columns: {sorted(missing)}. Found: {list(df.columns)}"
        )

    # Drop completely empty rows
    df = df.dropna(how="all")

    items = []
    for _, row in df.iterrows():
        risk_id = as_text(row.get("risk_id")).strip()
        title = as_text(row.get("title")).strip()

        # Skip blank lines
        if not risk_id and not title:
            continue

        item = {
            "id": risk_id,
            "description": title,
            "status": as_text(row.get("status")).strip(),
            "rating": normalise_rating(row.get("rating")),
            "ownerRole": as_text(row.get("owner_role")).strip(),
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
