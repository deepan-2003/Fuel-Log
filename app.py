"""Fleet Fuel Log - Streamlit app.
Run:  streamlit run app.py
"""
from __future__ import annotations

from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st

import analytics as an
import db
from excel_export import build_workbook, parse_report, sample_template

CUR = "₹"
DEFAULT_ORG = "O.P.R. Group of Educational Institutions"

st.set_page_config(page_title="Fleet Fuel Log", page_icon="⛽", layout="wide",
                   initial_sidebar_state="expanded")


def _ver(v: str) -> tuple:
    out = []
    for p in v.split(".")[:2]:
        out.append(int("".join(ch for ch in p if ch.isdigit()) or 0))
    return tuple(out)


# `use_container_width` was replaced by `width="stretch"` in newer Streamlit releases.
STRETCH = {"width": "stretch"} if _ver(st.__version__) >= (1, 50) else {"use_container_width": True}

db.init_db()

st.markdown("""
<style>
.block-container{padding-top:1.2rem;padding-bottom:2.5rem;max-width:1500px;}
.app-banner{background:linear-gradient(135deg,#1F3864 0%,#2E75B6 100%);padding:1.1rem 1.6rem;
  border-radius:14px;margin-bottom:1.1rem;box-shadow:0 4px 14px rgba(31,56,100,.18);}
.app-banner h1{margin:0;font-size:1.65rem;font-weight:700;color:#fff;letter-spacing:.2px;}
.app-banner p{margin:.25rem 0 0;color:#DCE8F7;font-size:.92rem;}
div[data-testid="stMetric"]{background:#fff;border:1px solid #E3E8F0;border-radius:12px;
  padding:.85rem 1rem;box-shadow:0 1px 3px rgba(16,24,40,.06);}
div[data-testid="stMetricLabel"] p{color:#5B6B84;font-size:.82rem;font-weight:600;}
div[data-testid="stMetricValue"]{color:#1F3864;font-size:1.55rem;font-weight:700;}
.sec{font-size:1.05rem;font-weight:700;color:#1F3864;margin:1.2rem 0 .4rem;
  padding-left:.6rem;border-left:4px solid #2E75B6;}
.chip{display:inline-block;background:#E8F1FB;color:#1F3864;border-radius:999px;
  padding:.15rem .7rem;font-size:.8rem;font-weight:600;margin-right:.35rem;}
button[data-baseweb="tab"]{font-weight:600;}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------------ helpers --
def section(title: str) -> None:
    st.markdown(f'<div class="sec">{title}</div>', unsafe_allow_html=True)


def money(x, dec=0) -> str:
    return "-" if x is None or pd.isna(x) else f"{CUR}{x:,.{dec}f}"


def num(x, dec=0) -> str:
    return "-" if x is None or pd.isna(x) else f"{x:,.{dec}f}"


def _blank(x):
    return "" if x is None or (not isinstance(x, str) and pd.isna(x)) else x


def _norm(v):
    """Normalise a cell so edited and original values compare cleanly."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.date()
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, float):
        return round(v, 4)
    return v


# --------------------------------------------------------- data + sidebar ---
vehicles = db.vehicles_df()
all_entries = db.entries_df()
reg_by_id = dict(zip(vehicles["id"], vehicles["reg_no"]))

with st.sidebar:
    st.markdown("## ⛽ Fleet Fuel Log")
    st.caption("Filters apply to the Dashboard, Manage Entries and Excel export.")
    sel_ids = st.multiselect("Vehicles", options=list(vehicles["id"]),
                             default=list(vehicles["id"]),
                             format_func=lambda i: reg_by_id.get(i, str(i)),
                             placeholder="All vehicles")
    period_choice = st.selectbox("Period", an.PERIOD_CHOICES)
    month_key, custom = None, (None, None)
    if period_choice == "Specific month":
        keys = sorted(all_entries["month_key"].unique(), reverse=True) if len(all_entries) else []
        if keys:
            month_key = st.selectbox("Month", keys, format_func=an.month_label)
        else:
            st.caption("No entries yet.")
    elif period_choice == "Custom range":
        c1 = st.date_input("From", value=date.today().replace(day=1), format="DD/MM/YYYY")
        c2 = st.date_input("To", value=date.today(), format="DD/MM/YYYY")
        custom = (c1, c2)
    start, end = an.resolve_period(period_choice, date.today(), custom, month_key)

    st.divider()
    org_name = st.text_input("Organisation name (for Excel report)",
                             value=db.get_setting("org_name", DEFAULT_ORG))
    if org_name != db.get_setting("org_name", DEFAULT_ORG):
        db.set_setting("org_name", org_name)

df = db.entries_df(vehicle_ids=sel_ids or None, start=start, end=end)
if len(df):
    period_label = f"{df['entry_date'].min():%d %b %Y} - {df['entry_date'].max():%d %b %Y}"
else:
    period_label = "No entries"
scope_vehicles = len(sel_ids) if sel_ids else len(vehicles)

st.markdown(
    f'<div class="app-banner"><h1>⛽ Fleet Fuel Log</h1>'
    f'<p>{org_name} &nbsp;•&nbsp; {scope_vehicles} vehicle(s) &nbsp;•&nbsp; {len(df)} fill(s) &nbsp;•&nbsp; {period_label}</p></div>',
    unsafe_allow_html=True)

if vehicles.empty:
    st.info("👋 Welcome! Add your first vehicle in the **🚘 Vehicles** tab, or load sample data to explore.")
    if st.button("Load sample data", type="secondary"):
        db.load_demo_data()
        st.rerun()

tab_dash, tab_new, tab_manage, tab_veh, tab_xl = st.tabs(
    ["📊 Dashboard", "➕ New Entry", "✏️ Manage Entries", "🚘 Vehicles", "📥 Export / Import"])


# ================================================================ DASHBOARD ==
with tab_dash:
    if df.empty:
        st.info("No fuel entries for the selected vehicles and period.")
    else:
        t = an.totals(df)
        k = st.columns(6)
        k[0].metric("Total KM", num(t["total_km"]))
        k[1].metric("Fuel filled", f"{num(t['fuel_litres'], 1)} L")
        k[2].metric("Fuel cost", money(t["amount"]))
        k[3].metric("Avg KM / L", num(t["mileage"], 2))
        k[4].metric("Cost / KM", money(t["cost_per_km"], 2))
        k[5].metric("No. of fills", num(t["fills"]))

        # FC reminders
        today = pd.Timestamp(date.today())
        shown = vehicles[vehicles["id"].isin(sel_ids)] if sel_ids else vehicles
        for _, v in shown.dropna(subset=["fc_date"]).iterrows():
            days = (v["fc_date"] - today).days
            if days < 0:
                st.error(f"🚨 FC for **{v['reg_no']}** expired {-days} day(s) ago ({v['fc_date']:%d-%m-%Y}).")
            elif days <= 30:
                st.warning(f"⚠️ FC for **{v['reg_no']}** is due in {days} day(s) ({v['fc_date']:%d-%m-%Y}).")

        section("Month-wise totals")
        view = st.radio("View", ["Vehicle-wise", "Combined"], horizontal=True, label_visibility="collapsed")
        ms = an.monthly_summary(df, by_vehicle=(view == "Vehicle-wise"))
        tbl = ms[["reg_no", "month_label", "fills", "total_km", "fuel_litres", "amount",
                  "mileage", "cost_per_km", "rate_per_litre"]].copy()
        tbl.loc[len(tbl)] = ["All vehicles" if view == "Vehicle-wise" else "All", "TOTAL", t["fills"],
                             t["total_km"], t["fuel_litres"], t["amount"], t["mileage"],
                             t["cost_per_km"], t["rate_per_litre"]]
        tbl.columns = ["Vehicle", "Month", "Fills", "Total KM", "Fuel (L)", f"Amount ({CUR})",
                       "Avg KM/L", f"Cost/KM ({CUR})", f"Rate/L ({CUR})"]
        styled = (tbl.style
                  .format({"Fills": "{:,.0f}", "Total KM": "{:,.0f}", "Fuel (L)": "{:,.2f}",
                           f"Amount ({CUR})": "{:,.2f}", "Avg KM/L": "{:.2f}",
                           f"Cost/KM ({CUR})": "{:.2f}", f"Rate/L ({CUR})": "{:.2f}"}, na_rep="-")
                  .set_properties(subset=pd.IndexSlice[tbl.index[-1], :],
                                  **{"font-weight": "700", "background-color": "#DDEBF7"}))
        st.dataframe(styled, hide_index=True, **STRETCH)

        section("Month × Vehicle matrix")
        metric = st.selectbox("Metric", list(an.METRICS), label_visibility="collapsed")
        mm = an.month_matrix(df, metric)
        mm.columns = [c if c == "Total" else an.month_label(c) for c in mm.columns]
        mm.index.name = "Vehicle"
        kind = an.METRICS[metric][1]
        fmt = {"int": "{:,.0f}", "dec": "{:,.2f}", "money": "{:,.0f}"}[kind]
        mstyled = (mm.style.format(fmt, na_rep="-")
                   .set_properties(subset=pd.IndexSlice[:, ["Total"]],
                                   **{"font-weight": "700", "background-color": "#DDEBF7"})
                   .set_properties(subset=pd.IndexSlice[["All vehicles"], :],
                                   **{"font-weight": "700", "background-color": "#DDEBF7"}))
        st.dataframe(mstyled, **STRETCH)

        section("Trends")
        by_v = an.monthly_summary(df, by_vehicle=True)
        c1, c2 = st.columns(2)
        with c1:
            st.caption(f"Monthly fuel cost ({CUR})")
            st.bar_chart(by_v.pivot(index="month_key", columns="reg_no", values="amount").fillna(0))
        with c2:
            st.caption("Average KM per litre by month")
            st.line_chart(by_v.pivot(index="month_key", columns="reg_no", values="mileage"))
        st.caption("Mileage of every fill (KM per litre)")
        st.line_chart(df.pivot_table(index="entry_date", columns="reg_no", values="mileage", aggfunc="mean"))

        section("Vehicle comparison")
        vs = an.vehicle_summary(df)
        vs_tbl = pd.DataFrame({
            "Vehicle": vs["reg_no"], "Fills": vs["fills"], "Total KM": vs["total_km"],
            "Fuel (L)": vs["fuel_litres"], f"Amount ({CUR})": vs["amount"],
            "Avg KM/L": vs["mileage"], f"Cost/KM ({CUR})": vs["cost_per_km"],
            "Last fill": vs["last_fill"].dt.strftime("%d-%m-%Y")})
        st.dataframe(vs_tbl.style.format({"Fills": "{:,.0f}", "Total KM": "{:,.0f}", "Fuel (L)": "{:,.2f}",
                                          f"Amount ({CUR})": "{:,.2f}", "Avg KM/L": "{:.2f}",
                                          f"Cost/KM ({CUR})": "{:.2f}"}, na_rep="-"),
                     hide_index=True, **STRETCH)

        valid = df.dropna(subset=["mileage"])
        if len(valid):
            best, worst = valid.loc[valid["mileage"].idxmax()], valid.loc[valid["mileage"].idxmin()]
            i1, i2 = st.columns(2)
            i1.success(f"🏆 Best mileage: **{best['mileage']:.2f} km/L** — {best['reg_no']} on {best['entry_date']:%d-%m-%Y}")
            i2.warning(f"📉 Lowest mileage: **{worst['mileage']:.2f} km/L** — {worst['reg_no']} on {worst['entry_date']:%d-%m-%Y}")


# ================================================================ NEW ENTRY ==
def _reset_entry_form():
    s = st.session_state
    s["ne_close"], s["ne_fuel"], s["ne_amount"] = 0, 0.0, 0.0
    s["ne_bill"], s["ne_remarks"] = "", ""


def _save_entry(vid: int, start_key: str):
    s = st.session_state
    try:
        bill_date = s["ne_date"] if s.get("ne_bill_same", True) else s["ne_bill_date"]
        db.add_entry(vid, s["ne_date"], s["ne_close"], s[start_key], s["ne_fuel"],
                     s["ne_bill"], bill_date, s["ne_amount"], s["ne_remarks"])
    except ValueError as exc:
        s["flash"] = ("error", str(exc))
        return
    s["flash"] = ("success", "Entry saved ✔")
    _reset_entry_form()


with tab_new:
    if vehicles.empty:
        st.info("Add a vehicle first (🚘 Vehicles tab).")
    else:
        flash = st.session_state.pop("flash", None)
        if flash:
            (st.success if flash[0] == "success" else st.error)(flash[1])

        active = vehicles[vehicles["active"]]
        pool = active if len(active) else vehicles
        left, right = st.columns([3, 2], gap="large")
        with left:
            section("Record a fill")
            vid = st.selectbox("Vehicle", list(pool["id"]), key="ne_vehicle",
                               format_func=lambda i: f"{reg_by_id[i]}  -  {pool.set_index('id').loc[i, 'name'] or 'vehicle'}")
            vrow = vehicles.set_index("id").loc[vid]
            last_close = db.last_closing_km(vid)
            start_key = f"ne_start_{vid}_{int(last_close)}"
            a, b = st.columns(2)
            a.date_input("Date", value=date.today(), key="ne_date", format="DD/MM/YYYY")
            b.number_input("Starting KM", min_value=0, step=1, value=int(last_close), key=start_key,
                           help="Pre-filled with the last closing KM of this vehicle.")
            a, b = st.columns(2)
            a.number_input("Closing KM", min_value=0, step=1, value=0, key="ne_close")
            b.number_input("Fuel filled (litres)", min_value=0.0, step=0.01, format="%.2f", key="ne_fuel")
            a, b = st.columns(2)
            a.number_input(f"Amount ({CUR})", min_value=0.0, step=0.01, format="%.2f", key="ne_amount")
            b.text_input("Bill No", key="ne_bill")
            st.checkbox("Bill date is the same as the fill date", value=True, key="ne_bill_same")
            if not st.session_state.get("ne_bill_same", True):
                st.date_input("Bill date", value=date.today(), key="ne_bill_date", format="DD/MM/YYYY")
            st.text_input("Remarks (optional)", key="ne_remarks")
            st.button("💾 Save entry", type="primary", on_click=_save_entry, args=(vid, start_key), **STRETCH)

        with right:
            section("Live preview")
            s = st.session_state
            km = (s.get("ne_close", 0) or 0) - (s.get(start_key, 0) or 0)
            fuel, amt = s.get("ne_fuel", 0.0) or 0.0, s.get("ne_amount", 0.0) or 0.0
            p1, p2 = st.columns(2)
            p1.metric("Total KM", num(km) if km > 0 else "-")
            p2.metric("Avg KM / L", num(km / fuel, 2) if km > 0 and fuel > 0 else "-")
            p3, p4 = st.columns(2)
            p3.metric("Rate / L", money(amt / fuel, 2) if fuel > 0 else "-")
            p4.metric("Cost / KM", money(amt / km, 2) if km > 0 and amt > 0 else "-")
            tank = float(vrow["tank_capacity"] or 0)
            if tank and fuel > tank:
                st.warning(f"Fuel ({fuel:g} L) is more than the tank capacity ({tank:g} L).")
            if s.get("ne_close", 0) and s.get("ne_close", 0) <= s.get(start_key, 0):
                st.warning("Closing KM must be greater than Starting KM.")
            if db.bill_exists(vid, s.get("ne_bill", "")):
                st.warning("This bill number is already recorded for this vehicle.")
            fc = vrow["fc_date"]
            if pd.notna(fc):
                st.caption(f"FC due: {fc:%d-%m-%Y}  |  Driver: {vrow['driver'] or '-'}")

            section("Recent entries")
            rec = all_entries[all_entries["vehicle_id"] == vid].sort_values(["entry_date", "id"], ascending=False).head(5)
            if len(rec):
                st.dataframe(pd.DataFrame({
                    "Date": rec["entry_date"].dt.strftime("%d-%m-%Y"), "KM": rec["total_km"],
                    "Litres": rec["fuel_litres"], "KM/L": rec["mileage"], f"{CUR}": rec["amount"]}
                ).style.format({"KM": "{:,.0f}", "Litres": "{:.2f}", "KM/L": "{:.2f}", f"{CUR}": "{:,.2f}"}),
                    hide_index=True, **STRETCH)
            else:
                st.caption("No entries yet for this vehicle.")


# ============================================================ MANAGE ENTRIES ==
EDIT_COLS = ["Date", "Closing KM", "Starting KM", "Fuel (L)", "Bill No", "Bill Date", "Amount", "Remarks"]

with tab_manage:
    if df.empty:
        st.info("No entries match the current filters.")
    else:
        section("Edit or delete entries")
        st.caption("Double-click a cell to edit. Tick **Delete** to remove a row. Then press **Save changes**.")
        view_df = df.sort_values(["entry_date", "id"], ascending=False).reset_index(drop=True)
        src = pd.DataFrame({
            "id": view_df["id"],
            "Delete": False,
            "Vehicle": view_df["reg_no"],
            "Date": view_df["entry_date"].dt.date,
            "Closing KM": view_df["closing_km"],
            "Starting KM": view_df["starting_km"],
            "Total KM": view_df["total_km"],
            "Fuel (L)": view_df["fuel_litres"],
            "Avg KM/L": view_df["mileage"].round(2),
            "Bill No": view_df["bill_no"].map(_blank),
            "Bill Date": [d.date() if pd.notna(d) else None for d in view_df["bill_date"]],
            "Amount": view_df["amount"],
            "Remarks": view_df["remarks"].map(_blank),
        })
        ver = st.session_state.setdefault("editor_ver", 0)
        edited = st.data_editor(
            src, key=f"editor_{ver}", hide_index=True, num_rows="fixed",
            disabled=["id", "Vehicle", "Total KM", "Avg KM/L"],
            column_config={
                "id": None,
                "Delete": st.column_config.CheckboxColumn("🗑 Delete", width="small"),
                "Date": st.column_config.DateColumn("Date", format="DD-MM-YYYY", required=True),
                "Bill Date": st.column_config.DateColumn("Bill Date", format="DD-MM-YYYY"),
                "Closing KM": st.column_config.NumberColumn("Closing KM", format="%d", min_value=0),
                "Starting KM": st.column_config.NumberColumn("Starting KM", format="%d", min_value=0),
                "Total KM": st.column_config.NumberColumn("Total KM", format="%d"),
                "Fuel (L)": st.column_config.NumberColumn("Fuel (L)", format="%.2f", min_value=0.0),
                "Avg KM/L": st.column_config.NumberColumn("Avg KM/L", format="%.2f"),
                "Amount": st.column_config.NumberColumn(f"Amount ({CUR})", format="%.2f", min_value=0.0),
            }, **STRETCH)

        n_del = int(edited["Delete"].sum())
        if st.button(f"💾 Save changes" + (f"  (delete {n_del})" if n_del else ""), type="primary"):
            upd, dele, errors = 0, [], []
            for i in range(len(src)):
                row_new, row_old = edited.iloc[i], src.iloc[i]
                eid = int(row_old["id"])
                if bool(row_new["Delete"]):
                    dele.append(eid)
                    continue
                if all(_norm(row_new[c]) == _norm(row_old[c]) for c in EDIT_COLS):
                    continue
                try:
                    db.update_entry(eid, row_new["Date"], row_new["Closing KM"], row_new["Starting KM"],
                                    row_new["Fuel (L)"], _blank(row_new["Bill No"]), row_new["Bill Date"],
                                    row_new["Amount"], _blank(row_new["Remarks"]))
                    upd += 1
                except Exception as exc:
                    errors.append(f"{row_old['Vehicle']} {row_old['Date']:%d-%m-%Y}: {exc}")
            db.delete_entries(dele)
            for msg in errors:
                st.error(msg)
            if upd or dele:
                st.toast(f"Updated {upd}, deleted {len(dele)}", icon="✅")
            if not errors:
                st.session_state["editor_ver"] = ver + 1
                st.rerun()


# ================================================================== VEHICLES ==
def _delete_vehicle(vehicle_id: int) -> None:
    """Runs as a button callback, so it may reset the checkbox's session state."""
    db.delete_vehicle(vehicle_id)
    st.session_state["del_confirm"] = False


with tab_veh:
    section("Add a vehicle")
    with st.form("add_vehicle", clear_on_submit=True):
        a, b, c = st.columns(3)
        reg = a.text_input("Registration No *", placeholder="TN-21-F-5500")
        name = b.text_input("Description", placeholder="e.g. Para Medical Science bus")
        driver = c.text_input("Driver")
        a, b, c = st.columns(3)
        tank = a.number_input("Tank capacity (L)", min_value=0.0, value=90.0, step=1.0)
        has_fc = b.checkbox("Set FC due date")
        fc = c.date_input("FC due date", value=date.today(), format="DD/MM/YYYY")
        if st.form_submit_button("Add vehicle", type="primary"):
            try:
                db.add_vehicle(reg, name, driver, tank, fc if has_fc else None)
                st.success(f"Vehicle {reg.strip().upper()} added.")
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))

    if not vehicles.empty:
        section("Your vehicles")
        vsrc = pd.DataFrame({
            "id": vehicles["id"], "Registration No": vehicles["reg_no"], "Description": vehicles["name"],
            "Driver": vehicles["driver"], "Tank (L)": vehicles["tank_capacity"],
            "FC due": [d.date() if pd.notna(d) else None for d in vehicles["fc_date"]],
            "Active": vehicles["active"]})
        vver = st.session_state.setdefault("veh_ver", 0)
        vedit = st.data_editor(
            vsrc, key=f"veh_editor_{vver}", hide_index=True, num_rows="fixed", disabled=["id", "Registration No"],
            column_config={"id": None,
                           "Tank (L)": st.column_config.NumberColumn(format="%.0f", min_value=0.0),
                           "FC due": st.column_config.DateColumn("FC due", format="DD-MM-YYYY"),
                           "Active": st.column_config.CheckboxColumn("Active")}, **STRETCH)
        if st.button("💾 Save vehicle changes"):
            for i in range(len(vsrc)):
                r = vedit.iloc[i]
                db.update_vehicle(int(r["id"]), _blank(r["Description"]), _blank(r["Driver"]),
                                  r["Tank (L)"], r["FC due"], bool(r["Active"]))
            st.session_state["veh_ver"] = vver + 1
            st.toast("Vehicles updated", icon="✅")
            st.rerun()

        with st.expander("🗑 Delete a vehicle"):
            del_id = st.selectbox("Vehicle to delete", list(vehicles["id"]),
                                  format_func=lambda i: reg_by_id[i], key="del_vehicle")
            n_ent = int((all_entries["vehicle_id"] == del_id).sum())
            st.warning(f"This also deletes its {n_ent} fuel entr{'y' if n_ent == 1 else 'ies'}. This cannot be undone.")
            confirmed = st.checkbox("I understand, delete permanently", key="del_confirm")
            st.button("Delete vehicle", disabled=not confirmed, on_click=_delete_vehicle, args=(int(del_id),))


# ============================================================ EXPORT / IMPORT ==
@st.cache_data(show_spinner=False)
def _make_xlsx(entries: pd.DataFrame, veh: pd.DataFrame, org: str, period: str) -> bytes:
    return build_workbook(entries, veh, org, period)


with tab_xl:
    section("Export to Excel")
    if df.empty:
        st.info("Nothing to export for the current filters.")
    else:
        n_veh = df["reg_no"].nunique()
        st.markdown(
            f'<span class="chip">{n_veh} vehicle sheet(s)</span><span class="chip">Summary sheet</span>'
            f'<span class="chip">{df["month_key"].nunique()} month(s)</span><span class="chip">{period_label}</span>',
            unsafe_allow_html=True)
        st.caption("Each vehicle sheet has month-wise sub totals and a grand total. "
                   "The Summary sheet shows months as columns for every vehicle, plus a cost chart. "
                   "Totals are live Excel formulas (click *Enable Editing* if Excel opens it in Protected View).")
        xlsx = _make_xlsx(df, vehicles, org_name, period_label)
        st.download_button("⬇️ Download Excel report", data=xlsx,
                           file_name=f"Fuel_Log_{date.today():%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")

    section("Restore / import data")
    st.caption("Upload an **Excel report exported from this app** to restore everything (vehicles and all entries), "
               "or a simple sheet with columns: " + ", ".join(db.IMPORT_COLUMNS) +
               ". Entries already present (same vehicle, date and closing KM) are skipped, "
               "so it is safe to upload the same file twice.")
    st.download_button("Download import template (same layout as the export)", data=sample_template(),
                       file_name="fuel_log_import_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    up = st.file_uploader("Upload file", type=["xlsx", "csv"])
    if up is not None:
        try:
            infos = []
            if up.name.lower().endswith(".csv"):
                imp = db.standardise_import(pd.read_csv(up))
            else:
                imp, infos = parse_report(up)
                if imp.empty:  # not an export-layout file: treat as a flat table
                    up.seek(0)
                    imp = db.standardise_import(pd.read_excel(up))
            if infos:
                st.success(f"Recognised an exported report: {len(infos)} vehicle sheet(s), {len(imp)} entr"
                           f"{'y' if len(imp) == 1 else 'ies'}.")
            st.dataframe(imp.head(20), hide_index=True, **STRETCH)
            st.caption(f"{len(imp)} row(s) found" + (" (first 20 shown)." if len(imp) > 20 else "."))
            create = st.checkbox("Create vehicles that don't exist yet", value=True)
            if st.button(f"Import {len(imp)} row(s)", type="primary"):
                if create and infos:
                    db.ensure_vehicles(infos)
                ins, skip, msgs = db.import_entries(imp, create)
                st.success(f"Imported {ins} row(s); skipped {skip}.")
                for m in msgs[:25]:
                    st.caption("• " + m)
                if ins:
                    st.rerun()
        except Exception as exc:
            st.error(f"Could not read the file: {exc}")

    section("Backup")
    st.download_button("Download data backup (.json)", data=db.backup_bytes(),
                       file_name=f"fuel_log_backup_{date.today():%Y%m%d}.json",
                       mime="application/json")
