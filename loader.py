import os
import pandas as pd
from pathlib import Path

REQUIRED_COLUMNS = ["date", "units_produced", "defects", "downtime_hours", "line_id", "shift"]


def load_file(path: str) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {path}")
    ext = p.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path, parse_dates=["date"])
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path, parse_dates=["date"])
    else:
        raise ValueError(f"Unsupported file format '{ext}'. Expected .csv, .xlsx, or .xls")


def validate_columns(df: pd.DataFrame, required: list) -> None:
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")


def summarize(df: pd.DataFrame) -> dict:
    numeric_cols = ["units_produced", "defects", "downtime_hours"]
    stats = {}
    for col in numeric_cols:
        stats[col] = {
            "mean": float(round(df[col].mean(), 2)),
            "min": float(round(df[col].min(), 2)),
            "max": float(round(df[col].max(), 2)),
            "null_count": int(df[col].isnull().sum()),
        }

    total_units = df["units_produced"].sum()
    total_defects = df["defects"].sum()
    defect_rate = round((total_defects / total_units * 100) if total_units > 0 else 0, 2)

    dates = df["date"].dropna()
    date_range = {
        "start": str(dates.min().date()) if len(dates) > 0 else "N/A",
        "end": str(dates.max().date()) if len(dates) > 0 else "N/A",
    }

    return {
        "row_count": int(len(df)),
        "date_range": date_range,
        "stats": stats,
        "defect_rate": defect_rate,
        "lines": sorted([str(x) for x in df["line_id"].unique().tolist()]),
        "shifts": sorted([str(x) for x in df["shift"].unique().tolist()]),
    }


if __name__ == "__main__":
    import json

    base = Path(__file__).parent / "samples"

    for filename in ["operations_sample.csv", "operations_sample.xlsx"]:
        path = base / filename
        if not path.exists():
            print(f"Skipping {filename} (not found)")
            continue
        print(f"\n--- {filename} ---")
        df = load_file(str(path))
        validate_columns(df, REQUIRED_COLUMNS)
        summary = summarize(df)
        print(json.dumps(summary, indent=2))
