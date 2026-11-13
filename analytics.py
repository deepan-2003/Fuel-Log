"""Pure-pandas helpers: derived columns, monthly roll-ups, period handling."""
from __future__ import annotations

import calendar
from datetime import date, timedelta

import pandas as pd

METRICS = {
    "Fuel cost (₹)": ("amount", "money"),
    "Total KM": ("total_km", "int"),
    "Fuel filled (L)": ("fuel_litres", "dec"),
    "No. of fills": ("fills", "int"),
    "Avg KM per litre": ("mileage", "dec"),
    "Cost per KM (₹)": ("cost_per_km", "dec"),
}

PERIOD_CHOICES = [
    "All time",
    "This month",
    "Last month",
    "Last 30 days",
    "This financial year (Apr-Mar)",
    "Specific month",
    "Custom range",
]


# --------------------------------------------------------------------------- #
def _numeric(series: pd.Series) -> pd.Series:
    """Convert PostgreSQL Decimal/object values to pandas numeric values."""
    return pd.to_numeric(series, errors="coerce").astype("float64")


def _ratio(num: pd.Series, den: pd.Series) -> pd.Series:
    """num / den, NaN wherever the denominator is zero or missing."""
    numerator = _numeric(num)
    denominator = _numeric(den)
    return numerator / denominator.where(denominator > 0)


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    """Add total_km, mileage, cost columns and month keys to an entries frame.

    PostgreSQL NUMERIC columns are returned by psycopg2 as decimal.Decimal.
    Convert numeric columns to float64 before arithmetic so Decimal/float
    operations do not raise TypeError in pandas.
    """
    df = df.copy()

    for col in ("closing_km", "starting_km", "fuel_litres", "amount"):
        if col in df.columns:
            df[col] = _numeric(df[col])

    df["entry_date"] = pd.to_datetime(df["entry_date"], errors="coerce")
    df["bill_date"] = pd.to_datetime(df["bill_date"], errors="coerce")
    df["total_km"] = df["closing_km"] - df["starting_km"]
    df["mileage"] = _ratio(df["total_km"], df["fuel_litres"])
    df["rate_per_litre"] = _ratio(df["amount"], df["fuel_litres"])
    df["cost_per_km"] = _ratio(df["amount"], df["total_km"])
    df["month_key"] = df["entry_date"].dt.strftime("%Y-%m")
    df["month_label"] = df["entry_date"].dt.strftime("%b %Y")
    return df


def month_label(key: str) -> str:
    return pd.Timestamp(key + "-01").strftime("%b %Y")


def totals(df: pd.DataFrame) -> dict:
    km = float(df["total_km"].sum()) if len(df) else 0.0
    fuel = float(df["fuel_litres"].sum()) if len(df) else 0.0
    amt = float(df["amount"].sum()) if len(df) else 0.0
    return {
        "fills": int(len(df)),
        "total_km": km,
        "fuel_litres": fuel,
        "amount": amt,
        "mileage": km / fuel if fuel > 0 else None,
        "cost_per_km": amt / km if km > 0 else None,
        "rate_per_litre": amt / fuel if fuel > 0 else None,
    }


def monthly_summary(df: pd.DataFrame, by_vehicle: bool = True) -> pd.DataFrame:
    cols = ["reg_no", "month_key", "month_label", "fills", "total_km",
            "fuel_litres", "amount", "mileage", "cost_per_km", "rate_per_litre"]
    if df.empty:
        return pd.DataFrame(columns=cols)
    keys = (["reg_no"] if by_vehicle else []) + ["month_key", "month_label"]
    g = df.groupby(keys, as_index=False).agg(
        fills=("id", "count"),
        total_km=("total_km", "sum"),
        fuel_litres=("fuel_litres", "sum"),
        amount=("amount", "sum"),
    )
    g["mileage"] = _ratio(g["total_km"], g["fuel_litres"])
    g["cost_per_km"] = _ratio(g["amount"], g["total_km"])
    g["rate_per_litre"] = _ratio(g["amount"], g["fuel_litres"])
    if not by_vehicle:
        g.insert(0, "reg_no", "All vehicles")
    return g.sort_values(["reg_no", "month_key"]).reset_index(drop=True)


def vehicle_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["reg_no", "fills", "total_km", "fuel_litres",
                                     "amount", "mileage", "cost_per_km", "last_fill"])
    g = df.groupby("reg_no", as_index=False).agg(
        fills=("id", "count"),
        total_km=("total_km", "sum"),
        fuel_litres=("fuel_litres", "sum"),
        amount=("amount", "sum"),
        last_fill=("entry_date", "max"),
    )
    g["mileage"] = _ratio(g["total_km"], g["fuel_litres"])
    g["cost_per_km"] = _ratio(g["amount"], g["total_km"])
    return g.sort_values("reg_no").reset_index(drop=True)


def month_keys_between(df: pd.DataFrame) -> list[str]:
    """Continuous list of YYYY-MM keys from the first to the last entry."""
    if df.empty:
        return []
    lo, hi = df["entry_date"].min(), df["entry_date"].max()
    return [str(p) for p in pd.period_range(lo, hi, freq="M")]


def _sum_matrix(d: pd.DataFrame, col: str, keys: list[str]) -> pd.DataFrame:
    p = (d.pivot_table(index="reg_no", columns="month_key", values=col,
                       aggfunc="sum", fill_value=0)
          .reindex(columns=keys, fill_value=0))
    p["Total"] = p.sum(axis=1)
    p.loc["All vehicles"] = p.sum(axis=0)
    return p


def month_matrix(df: pd.DataFrame, metric: str) -> pd.DataFrame:
    """Vehicles (+ 'All vehicles') as rows, months (+ 'Total') as columns."""
    if df.empty:
        return pd.DataFrame()
    d = df.copy()
    d["fills"] = 1
    keys = month_keys_between(d)
    field = METRICS[metric][0]
    if field in ("fills", "total_km", "fuel_litres", "amount"):
        return _sum_matrix(d, field, keys)
    km = _sum_matrix(d, "total_km", keys)
    if field == "mileage":
        fuel = _sum_matrix(d, "fuel_litres", keys)
        return km / fuel.where(fuel > 0)
    amount = _sum_matrix(d, "amount", keys)
    return amount / km.where(km > 0)


# --------------------------------------------------------------------------- #
def resolve_period(choice: str, today: date | None = None, custom=(None, None),
                   month_key: str | None = None):
    """Translate a period choice into an inclusive (start, end) date pair."""
    today = today or date.today()

    def month_bounds(y: int, m: int):
        return date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])

    if choice == "This month":
        return month_bounds(today.year, today.month)
    if choice == "Last month":
        first = today.replace(day=1) - timedelta(days=1)
        return month_bounds(first.year, first.month)
    if choice == "Last 30 days":
        return today - timedelta(days=29), today
    if choice.startswith("This financial year"):
        y = today.year if today.month >= 4 else today.year - 1
        return date(y, 4, 1), date(y + 1, 3, 31)
    if choice == "Specific month" and month_key:
        y, m = (int(x) for x in month_key.split("-"))
        return month_bounds(y, m)
    if choice == "Custom range":
        return custom
    return None, None
