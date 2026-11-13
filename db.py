"""SQLite persistence for vehicles and fuel entries."""
from __future__ import annotations

import os
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

import analytics

DB_PATH = Path(os.environ.get("FUEL_LOG_DB", Path(__file__).with_name("fuel_log.db")))

SCHEMA = """
CREATE TABLE IF NOT EXISTS vehicles(
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    reg_no        TEXT NOT NULL UNIQUE,
    name          TEXT DEFAULT '',
    driver        TEXT DEFAULT '',
    tank_capacity REAL DEFAULT 0,
    fc_date       TEXT,
    active        INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS fuel_entries(
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    vehicle_id  INTEGER NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE,
    entry_date  TEXT NOT NULL,
    closing_km  REAL NOT NULL,
    starting_km REAL NOT NULL,
    fuel_litres REAL NOT NULL,
    bill_no     TEXT DEFAULT '',
    bill_date   TEXT,
    amount      REAL NOT NULL DEFAULT 0,
    remarks     TEXT DEFAULT '',
    created_at  TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_entries_vehicle_date ON fuel_entries(vehicle_id, entry_date);
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
"""


@contextmanager
def connect():
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA foreign_keys = ON")
    try:
        yield con
        con.commit()
    finally:
        con.close()


def init_db() -> None:
    with connect() as con:
        con.executescript(SCHEMA)


# ------------------------------------------------------------------ helpers --
def _iso(d):
    """Any date-like value -> 'YYYY-MM-DD' (or None when empty)."""
    if d is None:
        return None
    if isinstance(d, str):
        d = d.strip()
        if not d:
            return None
        return pd.to_datetime(d, dayfirst=True).date().isoformat()
    if pd.isna(d):
        return None
    if isinstance(d, (datetime, pd.Timestamp)):
        return d.date().isoformat()
    if isinstance(d, date):
        return d.isoformat()
    raise ValueError(f"Unrecognised date: {d!r}")


def validate_entry(closing, starting, fuel, amount) -> None:
    if closing is None or starting is None or pd.isna(closing) or pd.isna(starting):
        raise ValueError("Closing and starting KM are required.")
    if starting < 0:
        raise ValueError("Starting KM cannot be negative.")
    if closing <= starting:
        raise ValueError("Closing KM must be greater than Starting KM.")
    if fuel is None or pd.isna(fuel) or fuel <= 0:
        raise ValueError("Fuel filled must be greater than 0 litres.")
    if amount is None or pd.isna(amount) or amount < 0:
        raise ValueError("Amount cannot be negative.")


# ----------------------------------------------------------------- settings --
def get_setting(key: str, default: str = "") -> str:
    with connect() as con:
        row = con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_setting(key: str, value: str) -> None:
    with connect() as con:
        con.execute("INSERT INTO settings(key,value) VALUES(?,?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))


# ----------------------------------------------------------------- vehicles --
def vehicles_df(active_only: bool = False) -> pd.DataFrame:
    q = "SELECT id, reg_no, name, driver, tank_capacity, fc_date, active FROM vehicles"
    if active_only:
        q += " WHERE active=1"
    q += " ORDER BY reg_no"
    with connect() as con:
        df = pd.read_sql_query(q, con)
    df["fc_date"] = pd.to_datetime(df["fc_date"], errors="coerce")
    df["active"] = df["active"].astype(bool)
    return df


def _tank(v) -> float:
    return 0.0 if v is None or pd.isna(v) else float(v)


def add_vehicle(reg_no: str, name: str = "", driver: str = "",
                tank_capacity: float = 0, fc_date=None) -> int:
    reg_no = (reg_no or "").strip().upper()
    if not reg_no:
        raise ValueError("Vehicle registration number is required.")
    try:
        with connect() as con:
            cur = con.execute(
                "INSERT INTO vehicles(reg_no,name,driver,tank_capacity,fc_date) VALUES(?,?,?,?,?)",
                (reg_no, (name or "").strip(), (driver or "").strip(),
                 _tank(tank_capacity), _iso(fc_date)))
            return cur.lastrowid
    except sqlite3.IntegrityError:
        raise ValueError(f"Vehicle {reg_no} already exists.") from None


def update_vehicle(vehicle_id: int, name: str, driver: str, tank_capacity: float,
                   fc_date, active: bool) -> None:
    with connect() as con:
        con.execute(
            "UPDATE vehicles SET name=?, driver=?, tank_capacity=?, fc_date=?, active=? WHERE id=?",
            ((name or "").strip(), (driver or "").strip(), _tank(tank_capacity),
             _iso(fc_date), 1 if active else 0, int(vehicle_id)))


def delete_vehicle(vehicle_id: int) -> None:
    with connect() as con:
        con.execute("DELETE FROM vehicles WHERE id=?", (int(vehicle_id),))


def vehicle_id_by_reg(reg_no: str):
    with connect() as con:
        row = con.execute("SELECT id FROM vehicles WHERE reg_no=?",
                          ((reg_no or "").strip().upper(),)).fetchone()
    return row[0] if row else None


# ------------------------------------------------------------------ entries --
def add_entry(vehicle_id, entry_date, closing_km, starting_km, fuel_litres,
              bill_no="", bill_date=None, amount=0.0, remarks="") -> int:
    validate_entry(closing_km, starting_km, fuel_litres, amount)
    with connect() as con:
        cur = con.execute(
            """INSERT INTO fuel_entries(vehicle_id, entry_date, closing_km, starting_km,
               fuel_litres, bill_no, bill_date, amount, remarks) VALUES(?,?,?,?,?,?,?,?,?)""",
            (int(vehicle_id), _iso(entry_date), float(closing_km), float(starting_km),
             float(fuel_litres), (bill_no or "").strip(), _iso(bill_date),
             float(amount), (remarks or "").strip()))
        return cur.lastrowid


def update_entry(entry_id, entry_date, closing_km, starting_km, fuel_litres,
                 bill_no, bill_date, amount, remarks) -> None:
    validate_entry(closing_km, starting_km, fuel_litres, amount)
    with connect() as con:
        con.execute(
            """UPDATE fuel_entries SET entry_date=?, closing_km=?, starting_km=?, fuel_litres=?,
               bill_no=?, bill_date=?, amount=?, remarks=? WHERE id=?""",
            (_iso(entry_date), float(closing_km), float(starting_km), float(fuel_litres),
             (bill_no or "").strip(), _iso(bill_date), float(amount),
             (remarks or "").strip(), int(entry_id)))


def delete_entries(ids) -> None:
    ids = [int(i) for i in ids]
    if not ids:
        return
    with connect() as con:
        con.execute(f"DELETE FROM fuel_entries WHERE id IN ({','.join('?' * len(ids))})", ids)


def entries_df(vehicle_ids=None, start=None, end=None) -> pd.DataFrame:
    q = """SELECT e.id, e.vehicle_id, v.reg_no, e.entry_date, e.closing_km, e.starting_km,
                  e.fuel_litres, e.bill_no, e.bill_date, e.amount, e.remarks
           FROM fuel_entries e JOIN vehicles v ON v.id = e.vehicle_id WHERE 1=1"""
    params: list = []
    if vehicle_ids:
        q += f" AND e.vehicle_id IN ({','.join('?' * len(vehicle_ids))})"
        params += [int(i) for i in vehicle_ids]
    if start:
        q += " AND e.entry_date >= ?"
        params.append(_iso(start))
    if end:
        q += " AND e.entry_date <= ?"
        params.append(_iso(end))
    q += " ORDER BY e.entry_date, e.id"
    with connect() as con:
        df = pd.read_sql_query(q, con, params=params)
    return analytics.enrich(df)


def last_closing_km(vehicle_id) -> float:
    with connect() as con:
        row = con.execute(
            "SELECT closing_km FROM fuel_entries WHERE vehicle_id=? "
            "ORDER BY entry_date DESC, id DESC LIMIT 1", (int(vehicle_id),)).fetchone()
    return float(row[0]) if row else 0.0


def bill_exists(vehicle_id, bill_no: str, exclude_id=None) -> bool:
    bill_no = (bill_no or "").strip()
    if not bill_no:
        return False
    q = "SELECT 1 FROM fuel_entries WHERE vehicle_id=? AND bill_no=?"
    params: list = [int(vehicle_id), bill_no]
    if exclude_id is not None:
        q += " AND id<>?"
        params.append(int(exclude_id))
    with connect() as con:
        return con.execute(q, params).fetchone() is not None


def backup_bytes() -> bytes:
    return DB_PATH.read_bytes() if DB_PATH.exists() else b""


# ------------------------------------------------------------------- import --
IMPORT_COLUMNS = ["Vehicle", "Date", "Closing KM", "Starting KM", "Fuel (L)",
                  "Bill No", "Bill Date", "Amount", "Remarks"]
_ALIASES = {
    "Vehicle": {"vehicle", "vehicleno", "regno", "registration", "regnumber", "vehiclenumber"},
    "Date": {"date", "entrydate"},
    "Closing KM": {"closingkm", "closing", "closingreading"},
    "Starting KM": {"startingkm", "starting", "openingkm", "startkm"},
    "Fuel (L)": {"fuell", "fuel", "fuelfilled", "fuelfilledl", "litres", "liters", "fuelquantity"},
    "Bill No": {"billno", "billnumber", "bill"},
    "Bill Date": {"billdate"},
    "Amount": {"amount", "amountrs", "amountinr", "cost"},
    "Remarks": {"remarks", "remark", "notes"},
}


def _norm_header(h) -> str:
    return "".join(ch for ch in str(h).lower() if ch.isalnum())


def standardise_import(raw: pd.DataFrame) -> pd.DataFrame:
    """Rename the uploaded sheet's headers to IMPORT_COLUMNS; raise if required ones are missing."""
    lookup = {a: std for std, names in _ALIASES.items() for a in names}
    rename = {}
    for col in raw.columns:
        std = lookup.get(_norm_header(col))
        if std and std not in rename.values():
            rename[col] = std
    df = raw.rename(columns=rename)
    missing = [c for c in ["Vehicle", "Date", "Closing KM", "Starting KM", "Fuel (L)", "Amount"]
               if c not in df.columns]
    if missing:
        raise ValueError("Missing column(s): " + ", ".join(missing))
    for c in IMPORT_COLUMNS:
        if c not in df.columns:
            df[c] = None
    return df[IMPORT_COLUMNS].dropna(how="all")


def ensure_vehicles(infos) -> int:
    """Create any vehicle (with driver, tank, FC date) that does not exist yet."""
    made = 0
    for v in infos:
        if vehicle_id_by_reg(v["reg_no"]) is None:
            add_vehicle(v["reg_no"], v.get("name", ""), v.get("driver", ""),
                        v.get("tank_capacity", 0), v.get("fc_date"))
            made += 1
    return made


def import_entries(df: pd.DataFrame, create_missing: bool = True):
    """Insert rows from a standardised frame. Returns (inserted, skipped, messages)."""
    inserted, skipped, msgs = 0, 0, []
    for i, row in df.iterrows():
        line = i + 2  # header is line 1 in the source sheet
        try:
            reg = str(row["Vehicle"]).strip().upper()
            vid = vehicle_id_by_reg(reg)
            if vid is None:
                if not create_missing or not reg or reg == "NAN":
                    raise ValueError(f"unknown vehicle '{reg}'")
                vid = add_vehicle(reg)
            bill = "" if pd.isna(row["Bill No"]) else str(row["Bill No"]).strip()
            if bill.endswith(".0"):
                bill = bill[:-2]
            remarks = "" if pd.isna(row["Remarks"]) else str(row["Remarks"])
            if _is_duplicate(vid, row):
                skipped += 1
                msgs.append(f"Line {line}: skipped (same vehicle, date and closing KM already exist)")
                continue
            add_entry(vid, row["Date"], float(row["Closing KM"]), float(row["Starting KM"]),
                      float(row["Fuel (L)"]), bill, row["Bill Date"],
                      float(row["Amount"]) if not pd.isna(row["Amount"]) else 0.0, remarks)
            inserted += 1
        except Exception as exc:  # keep importing the remaining rows
            skipped += 1
            msgs.append(f"Line {line}: {exc}")
    return inserted, skipped, msgs


def _is_duplicate(vehicle_id, row) -> bool:
    with connect() as con:
        return con.execute(
            "SELECT 1 FROM fuel_entries WHERE vehicle_id=? AND entry_date=? AND closing_km=?",
            (int(vehicle_id), _iso(row["Date"]), float(row["Closing KM"]))).fetchone() is not None


# --------------------------------------------------------------------- demo --
def load_demo_data() -> None:
    """Two sample vehicles with ~4 months of fills so the dashboard has something to show."""
    demo = [
        ("DEMO-BUS-1", "Sample staff bus", "Driver A", 90, 8.2, 70, 585),
        ("DEMO-BUS-2", "Sample college van", "Driver B", 60, 10.5, 45, 470),
    ]
    start = date.today().replace(day=1) - timedelta(days=120)
    for reg, name, drv, tank, kmpl, litres, km in demo:
        if vehicle_id_by_reg(reg):
            continue
        vid = add_vehicle(reg, name, drv, tank, date.today() + timedelta(days=200))
        odo = 40000.0
        for k in range(16):
            wobble = ((k * 7) % 5 - 2) * 0.25
            fuel = round(litres + ((k * 3) % 5 - 2) * 1.6, 2)
            dist = round(fuel * (kmpl + wobble))
            d = start + timedelta(days=k * 7)
            add_entry(vid, d, odo + dist, odo, fuel, f"B{1000 + k}{vid}", d,
                      round(fuel * (97.5 + k * 0.3), 2), "")
            odo += dist
