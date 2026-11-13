"""Professional Excel report: a Summary sheet (months as columns) + one sheet per vehicle
with month-wise sub totals and a grand total. Totals are live Excel formulas."""
from __future__ import annotations

import io
import re
from datetime import date, datetime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

import analytics as an

FONT = "Arial"
NAVY, HEAD, BLUE = "1F3864", "1F4E78", "2E75B6"
BAND, SUB, LIGHT, GRID = "DDEBF7", "BDD7EE", "F2F2F2", "BFBFBF"
CUR = "₹"

F_KM = '#,##0;-#,##0;"-"'
F_DEC = '#,##0.00;-#,##0.00;"-"'
F_MONEY = f'"{CUR}" #,##0.00;-"{CUR}" #,##0.00;"-"'
F_INT = '0;-0;"-"'
F_DATE = "dd-mm-yyyy"

_thin = Side(style="thin", color=GRID)
BOX = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
TOP_RULE = Border(left=_thin, right=_thin, bottom=_thin, top=Side(style="medium", color=NAVY))


def _fill(hex_: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def _put(ws, row, col, value=None, *, bold=False, italic=False, color="000000", size=10,
         fill=None, h="center", v="center", fmt=None, border=BOX, wrap=False):
    c = ws.cell(row, col)
    if value is not None:
        c.value = value
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if fill:
        c.fill = _fill(fill)
    if fmt:
        c.number_format = fmt
    if border:
        c.border = border
    return c


def _merge(ws, row, c1, c2, value=None, **kw):
    """Merge a horizontal range, styling every cell so borders/fills render fully."""
    for c in range(c1, c2 + 1):
        _put(ws, row, c, value if c == c1 else None, **kw)
    if c2 > c1:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)


def _banner(ws, ncol, org, subtitle, meta):
    _merge(ws, 1, 1, ncol, org, bold=True, size=16, color="FFFFFF", fill=NAVY, border=None)
    ws.row_dimensions[1].height = 30
    _merge(ws, 2, 1, ncol, subtitle, bold=True, size=12, color="FFFFFF", fill=BLUE, border=None)
    ws.row_dimensions[2].height = 22
    _merge(ws, 3, 1, ncol, meta, italic=True, size=9, color="595959", border=None, h="left")


def _safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]\:\*\?\/\\']", "-", name).strip()[:31] or "Vehicle"
    cand, n = base, 2
    while cand.lower() in used or cand.lower() == "summary":
        suffix = f" ({n})"
        cand = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(cand.lower())
    return cand


def _page_setup(ws, title_rows=None):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.size = 8
    if title_rows:
        ws.print_title_rows = title_rows


# --------------------------------------------------------------------------- #
VEH_HEADERS = ["Date", "Closing KM", "Starting KM", "Total KM", "Fuel Filled (L)",
               "Avg KM/L", "Bill No", "Bill Date", f"Amount ({CUR})", f"Cost / KM ({CUR})", "Remarks"]
VEH_WIDTHS = [13, 13, 13, 12, 14, 11, 16, 13, 16, 14, 28]


def _vehicle_sheet(wb, title, veh, g, org, period, generated):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = BLUE
    n = len(VEH_HEADERS)
    for i, w in enumerate(VEH_WIDTHS, 1):
        ws.column_dimensions[L(i)].width = w

    _banner(ws, n, org, f"Fuel Consumption Log  -  {veh['reg_no']}",
            f"Period: {period}     |     Generated: {generated:%d-%m-%Y %H:%M}")

    tank = float(veh.get("tank_capacity") or 0)
    fc = veh.get("fc_date")
    info = [
        (4, "Vehicle No.", veh["reg_no"], "Driver", veh.get("driver") or "-"),
        (5, "Description", veh.get("name") or "-", "Tank Capacity",
         f"{tank:g} L" if tank else "-"),
        (6, "FC Due", fc.strftime("%d-%m-%Y") if pd.notna(fc) else "-", "Period", period),
    ]
    for r, l1, v1, l2, v2 in info:
        _put(ws, r, 1, l1, bold=True, fill=LIGHT, h="left")
        _merge(ws, r, 2, 4, v1, h="left")
        _put(ws, r, 5, l2, bold=True, fill=LIGHT, h="left", wrap=True)
        _merge(ws, r, 6, n, v2, h="left")
    ws.row_dimensions[8].height = 34
    for i, h in enumerate(VEH_HEADERS, 1):
        _put(ws, 8, i, h, bold=True, color="FFFFFF", fill=HEAD, wrap=True)
    ws.freeze_panes = "A9"

    r = 9
    sub_rows: dict[str, int] = {}
    for mk, grp in g.groupby("month_key", sort=True):
        label = grp["month_label"].iloc[0]
        _merge(ws, r, 1, n, label, bold=True, color=NAVY, fill=BAND, h="left", size=11)
        ws.row_dimensions[r].height = 20
        r += 1
        first = r
        for _, e in grp.sort_values(["entry_date", "id"]).iterrows():
            bill_date = e["bill_date"].date() if pd.notna(e["bill_date"]) else None
            _put(ws, r, 1, e["entry_date"].date(), fmt=F_DATE)
            _put(ws, r, 2, float(e["closing_km"]), fmt="#,##0")
            _put(ws, r, 3, float(e["starting_km"]), fmt="#,##0")
            _put(ws, r, 4, f"=B{r}-C{r}", fmt="#,##0")
            _put(ws, r, 5, float(e["fuel_litres"]), fmt="0.00")
            _put(ws, r, 6, f"=IFERROR(D{r}/E{r},0)", fmt="0.00")
            _put(ws, r, 7, e["bill_no"] or None)
            _put(ws, r, 8, bill_date, fmt=F_DATE)
            _put(ws, r, 9, float(e["amount"]), fmt=F_MONEY, h="right")
            _put(ws, r, 10, f"=IFERROR(I{r}/D{r},0)", fmt=F_MONEY, h="right")
            _put(ws, r, 11, e["remarks"] or None, h="left", wrap=True)
            r += 1
        last = r - 1
        _merge(ws, r, 1, 3, f"Sub Total  -  {label}", bold=True, color=NAVY, fill=SUB,
               h="left", border=TOP_RULE)
        _put(ws, r, 4, f"=SUM(D{first}:D{last})", bold=True, fill=SUB, fmt="#,##0", border=TOP_RULE)
        _put(ws, r, 5, f"=SUM(E{first}:E{last})", bold=True, fill=SUB, fmt="0.00", border=TOP_RULE)
        _put(ws, r, 6, f"=IFERROR(D{r}/E{r},0)", bold=True, fill=SUB, fmt="0.00", border=TOP_RULE)
        _put(ws, r, 7, f"=COUNT(D{first}:D{last})", bold=True, fill=SUB, fmt='0" fills"', border=TOP_RULE)
        _put(ws, r, 8, None, fill=SUB, border=TOP_RULE)
        _put(ws, r, 9, f"=SUM(I{first}:I{last})", bold=True, fill=SUB, fmt=F_MONEY, h="right", border=TOP_RULE)
        _put(ws, r, 10, f"=IFERROR(I{r}/D{r},0)", bold=True, fill=SUB, fmt=F_MONEY, h="right", border=TOP_RULE)
        _put(ws, r, 11, None, fill=SUB, border=TOP_RULE)
        sub_rows[mk] = r
        r += 1
        ws.row_dimensions[r].height = 8
        r += 1

    refs = list(sub_rows.values())

    def tot(col):
        return "=SUM(" + ",".join(f"{col}{x}" for x in refs) + ")"

    _merge(ws, r, 1, 3, "GRAND TOTAL", bold=True, color="FFFFFF", fill=NAVY, h="left", size=11)
    ws.row_dimensions[r].height = 22
    _put(ws, r, 4, tot("D"), bold=True, color="FFFFFF", fill=NAVY, fmt="#,##0")
    _put(ws, r, 5, tot("E"), bold=True, color="FFFFFF", fill=NAVY, fmt="0.00")
    _put(ws, r, 6, f"=IFERROR(D{r}/E{r},0)", bold=True, color="FFFFFF", fill=NAVY, fmt="0.00")
    _put(ws, r, 7, tot("G"), bold=True, color="FFFFFF", fill=NAVY, fmt='0" fills"')
    _put(ws, r, 8, None, fill=NAVY)
    _put(ws, r, 9, tot("I"), bold=True, color="FFFFFF", fill=NAVY, fmt=F_MONEY, h="right")
    _put(ws, r, 10, f"=IFERROR(I{r}/D{r},0)", bold=True, color="FFFFFF", fill=NAVY, fmt=F_MONEY, h="right")
    _put(ws, r, 11, None, fill=NAVY)
    r += 2
    _merge(ws, r, 1, n, "Avg KM/L = Total KM / Fuel filled.   Cost / KM = Amount / Total KM.   "
                         "Sub totals and grand total are live formulas.",
           italic=True, size=8, color="7F7F7F", border=None, h="left")
    _page_setup(ws, title_rows="8:8")
    return sub_rows


# --------------------------------------------------------------------------- #
def _summary_sheet(ws, vehicles, sub_rows, months, org, period, generated):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    nm = len(months)
    ncol = 3 + nm
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 18
    for c in range(3, ncol):
        ws.column_dimensions[L(c)].width = 14
    ws.column_dimensions[L(ncol)].width = 17

    _banner(ws, ncol, org, "Fuel Consumption Summary  -  Month-wise by Vehicle",
            f"Period: {period}     |     Generated: {generated:%d-%m-%Y %H:%M}")
    hdr = 5
    ws.row_dimensions[hdr].height = 26
    _put(ws, hdr, 1, "Vehicle", bold=True, color="FFFFFF", fill=HEAD)
    _put(ws, hdr, 2, "Metric", bold=True, color="FFFFFF", fill=HEAD)
    for i, mk in enumerate(months):
        _put(ws, hdr, 3 + i, an.month_label(mk), bold=True, color="FFFFFF", fill=HEAD)
    _put(ws, hdr, ncol, "Total", bold=True, color="FFFFFF", fill=NAVY)
    ws.freeze_panes = ws.cell(hdr + 1, 3)

    metrics = [("No. of fills", F_INT, "G"), ("Total KM", F_KM, "D"), ("Fuel (L)", F_DEC, "E"),
               (f"Amount ({CUR})", F_MONEY, "I"), ("Avg KM/L", F_DEC, None), (f"Cost / KM ({CUR})", F_MONEY, None)]
    first_m, last_m = L(3), L(2 + nm)

    def write_block(r0, label, sources, fill_a, strong=False):
        """sources[(row_offset, col_idx)] -> formula body; ratio rows derive from same column."""
        fill_row = SUB if strong else None
        merge_end = r0 + len(metrics) - 1
        _put(ws, r0, 1, label, bold=True, color="FFFFFF" if strong else NAVY, fill=fill_a, wrap=True)
        for k in range(1, len(metrics)):
            _put(ws, r0 + k, 1, None, fill=fill_a)
        ws.merge_cells(start_row=r0, start_column=1, end_row=merge_end, end_column=1)
        for k, (name, fmt, _src) in enumerate(metrics):
            r = r0 + k
            _put(ws, r, 2, name, bold=strong, h="left", fill=fill_row, italic=k >= 4)
            for ci in range(3, ncol + 1):
                col = L(ci)
                if k < 4:
                    val = (f"=SUM({first_m}{r}:{last_m}{r})" if ci == ncol else sources(k, ci))
                elif k == 4:
                    val = f"=IFERROR({col}{r0 + 1}/{col}{r0 + 2},0)"
                else:
                    val = f"=IFERROR({col}{r0 + 3}/{col}{r0 + 1},0)"
                _put(ws, r, ci, val, fmt=fmt, bold=strong or ci == ncol, fill=fill_row or (BAND if ci == ncol else None),
                     h="right", italic=k >= 4)
        return r0

    r = hdr + 1
    blocks = []
    for veh in vehicles:
        rows = sub_rows[veh["reg_no"]]
        sheet = veh["sheet"]

        def src(k, ci, rows=rows, sheet=sheet):
            mk = months[ci - 3]
            if mk not in rows:
                return 0
            return f"='{sheet}'!{metrics[k][2]}{rows[mk]}"

        write_block(r, veh["reg_no"], src, BAND)
        blocks.append((veh["reg_no"], r))
        r += len(metrics)
        ws.row_dimensions[r].height = 6
        r += 1

    if len(blocks) > 1:
        def fleet(k, ci):
            col = L(ci)
            return "=SUM(" + ",".join(f"{col}{b + k}" for _, b in blocks) + ")"
        write_block(r, "ALL VEHICLES", fleet, NAVY, strong=True)
        r += len(metrics) + 1

    _merge(ws, r, 1, ncol, "Figures link to the sub totals on each vehicle sheet. "
                           "Avg KM/L = Total KM / Fuel; Cost / KM = Amount / Total KM.",
           italic=True, size=8, color="7F7F7F", border=None, h="left")
    r += 2

    # Monthly fuel-cost chart
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.title = f"Monthly Fuel Cost ({CUR}) by Vehicle"
    ch.roundedCorners = False
    ch.title.overlay = False
    ch.style = 10
    for reg, b in blocks:
        s = Series(Reference(ws, min_col=3, max_col=2 + nm, min_row=b + 3, max_row=b + 3), title=reg)
        ch.series.append(s)
    ch.set_categories(Reference(ws, min_col=3, max_col=2 + nm, min_row=hdr, max_row=hdr))
    ch.y_axis.title = f"Amount ({CUR})"
    ch.y_axis.title.overlay = False
    ch.y_axis.numFmt = "#,##0"
    ch.y_axis.majorGridlines = ch.y_axis.majorGridlines
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.height, ch.width = 9, max(18, min(34, 3 + nm * 3.2))
    ws.add_chart(ch, f"A{r}")
    _page_setup(ws)


# --------------------------------------------------------------------------- #
def build_workbook(entries: pd.DataFrame, vehicles: pd.DataFrame, org_name: str,
                   period_label: str) -> bytes:
    """entries must come from db.entries_df(); vehicles from db.vehicles_df()."""
    generated = datetime.now()
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"

    vmap = vehicles.set_index("reg_no").to_dict("index")
    used: set[str] = set()
    veh_list, sub_rows = [], {}
    for reg in sorted(entries["reg_no"].unique()):
        g = entries[entries["reg_no"] == reg]
        info = {"reg_no": reg, **vmap.get(reg, {})}
        info["sheet"] = _safe_sheet_name(reg, used)
        sub_rows[reg] = _vehicle_sheet(wb, info["sheet"], info, g, org_name, period_label, generated)
        veh_list.append(info)

    _summary_sheet(summary, veh_list, sub_rows, an.month_keys_between(entries),
                   org_name, period_label, generated)

    wb.properties.title = "Fuel Consumption Report"
    wb.properties.creator = org_name
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Reading a report back in (restore / re-import)
IMPORT_COLS = ["Vehicle", "Date", "Closing KM", "Starting KM", "Fuel (L)",
               "Bill No", "Bill Date", "Amount", "Remarks"]


def _clean(v):
    return "" if v is None or str(v).strip() in ("-", "None") else str(v).strip()


def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def parse_report(file):
    """Read a workbook produced by build_workbook().

    Returns (entries_frame, vehicle_infos). The frame uses the import column names;
    sub total / grand total / month-band rows are ignored. Returns (empty, []) when the
    file is not in the report layout.
    """
    wb = load_workbook(file, data_only=True)
    rows, infos = [], []
    for ws in wb.worksheets:
        if ws.title == "Summary" or str(ws["A8"].value or "").strip() != "Date":
            continue
        reg = _clean(ws["B4"].value).upper()
        if not reg:
            continue
        tank = re.search(r"[\d.]+", _clean(ws["F5"].value))
        infos.append({"reg_no": reg, "name": _clean(ws["B5"].value), "driver": _clean(ws["F4"].value),
                      "tank_capacity": float(tank.group()) if tank else 0.0,
                      "fc_date": _clean(ws["B6"].value) or None})
        for r in ws.iter_rows(min_row=9, max_col=11, values_only=True):
            d = r[0]
            if isinstance(d, (datetime, date)) and _is_num(r[1]) and _is_num(r[2]) and _is_num(r[4]):
                bill = r[6]
                if isinstance(bill, float) and bill.is_integer():
                    bill = int(bill)
                rows.append([reg, d, r[1], r[2], r[4], _clean(bill), r[7], r[8] if _is_num(r[8]) else 0, _clean(r[10])])
    return pd.DataFrame(rows, columns=IMPORT_COLS), infos


def sample_template() -> bytes:
    """A small workbook in exactly the export layout, usable as an import template."""
    d = lambda y, m, dd: pd.Timestamp(y, m, dd)
    raw = pd.DataFrame({
        "id": [1, 2, 3], "vehicle_id": [1, 1, 1], "reg_no": ["TN-21-F-5500"] * 3,
        "entry_date": [d(2026, 4, 7), d(2026, 4, 12), d(2026, 5, 20)],
        "closing_km": [49608, 50183, 50665], "starting_km": [49025, 49608, 50183],
        "fuel_litres": [71.0, 72.06, 58.62], "bill_no": ["45925", "45925", "46209"],
        "bill_date": [d(2026, 4, 7), d(2026, 4, 12), d(2026, 5, 20)],
        "amount": [6700.0, 6840.0, 5750.0], "remarks": ["", "", "Sample row - replace"]})
    entries = an.enrich(raw)
    veh = pd.DataFrame([{"reg_no": "TN-21-F-5500", "name": "Vehicle description", "driver": "Driver name",
                         "tank_capacity": 90.0, "fc_date": pd.Timestamp(2026, 8, 11)}])
    return build_workbook(entries, veh, "Your organisation", "Template")
