import json
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_XLSX = REPO_ROOT / "source" / "project_dashboard.xlsx"

RISKS_JSON = REPO_ROOT / "data" / "risks.json"
TQS_JSON = REPO_ROOT / "data" / "tqs.json"

RISK_SHEET_CANDIDATES = ["Risks", "Risk"]
TQ_SHEET_CANDIDATES = ["TQs", "TQ", "Tqs"]


def find_sheet_name(xls: pd.ExcelFile, candidates: list[str]) -> str | None:
    for name in candidates:
        if name in xls.sheet_names:
            return name
    return None


def as_text(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val)


def normalise_rating(x) -> str:
    s = as_text(x).strip()
    if not s:
        return ""
    s_low = s.lower()

    if s_low in {"red", "r", "high"}:
        return "Red"
    if s_low in {"amber", "orange", "a", "medium", "med"}:
        return "Amber"
    if s_low in {"green", "g", "low"}:
        return "Green"

    return s[:1].upper() + s[1:].lower()


def to_iso_date(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass

    if isinstance(val, pd.Timestamp):
        return val.date().isoformat()

    if hasattr(val, "date"):
        try:
            return val.date().isoformat()
        except Exception:
            pass

    return as_text(val).strip()


def read_sheet_df(sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(
        SOURCE_XLSX,
        sheet_name=sheet_name,
        engine="openpyxl",
        dtype=object,  # keep dates as dates; keep ids stable
    ).dropna(how="all")


def build_risks(df: pd.DataFrame) -> list[dict]:
    expected = {"risk_id", "title", "status", "rating", "due_date", "owner_role", "last_updated"}
    missing = expected - set(df.columns)
    if missing:
        raise ValueError(f"Risks sheet missing columns: {sorted(missing)}. Found: {list(df.columns)}")

    items: list[dict] = []
    for _, row in df.iterrows():
        risk_id = as_text(row.get("risk_id")).strip()
        title = as_text(row.get("title")).strip()

        if not risk_id and not title:
            continue

        items.append(
            {
                "id": risk_id,
                "description": title,
                "status": as_text(row.get("status")).strip(),
                "rating": normalise_rating(row.get("rating")),
                "ownerRole": as_text(row.get("owner_role")).strip(),
                "nextActionDate": to_iso_date(row.get("due_date")),
                "lastUpdatedRow": to_iso_date(row.get("last_updated")),
            }
        )
    return items


def build_tqs(df: pd.DataFrame) -> list[dict]:
    # Minimal TQ schema (matches what you have)
    expected = {"tq_id", "title", "status"}
    missing = expected - set(df.columns)
    if missing:
        raise ValueError(f"TQs sheet missing columns: {sorted(missing)}. Found: {list(df.columns)}")

    items: list[dict] = []
    for _, row in df.iterrows():
        tq_id = as_text(row.get("tq_id")).strip()
        title = as_text(row.get("title")).strip()

        if not tq_id and not title:
            continue

        items.append(
            {
                "id": tq_id,
                "title": title,
                "status": as_text(row.get("status")).strip(),
            }
        )
    return items


def write_json(path: Path, payload: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def main():
    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"Missing source workbook: {SOURCE_XLSX}")

    xls = pd.ExcelFile(SOURCE_XLSX, engine="openpyxl")

    now_awst = datetime.now(ZoneInfo("Australia/Perth")).strftime("%Y-%m-%d %H:%M AWST")

    # --- Risks ---
    risk_sheet = find_sheet_name(xls, RISK_SHEET_CANDIDATES)
    if not risk_sheet:
        raise ValueError(f"No risks sheet found. Tried {RISK_SHEET_CANDIDATES}. Found: {xls.sheet_names}")

    risks_df = read_sheet_df(risk_sheet)
    risks_items = build_risks(risks_df)
    write_json(RISKS_JSON, {"lastUpdated": now_awst, "items": risks_items})
    print(f"Wrote {RISKS_JSON} with {len(risks_items)} items from sheet '{risk_sheet}'")

    # --- TQs ---
    tq_sheet = find_sheet_name(xls, TQ_SHEET_CANDIDATES)
    if not tq_sheet:
        raise ValueError(f"No TQs sheet found. Tried {TQ_SHEET_CANDIDATES}. Found: {xls.sheet_names}")

    tqs_df = read_sheet_df(tq_sheet)
    tqs_items = build_tqs(tqs_df)
    write_json(TQS_JSON, {"lastUpdated": now_awst, "items": tqs_items})
    print(f"Wrote {TQS_JSON} with {len(tqs_items)} items from sheet '{tq_sheet}'")


if __name__ == "__main__":
    main()
