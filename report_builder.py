"""
ตรรกะสร้างรายงานตรวจครุภัณฑ์ประจำปี ใช้ร่วมกันทั้งจาก CLI (rpt_asset_annual_inspection.py)
และเว็บแอป (app.py)

หมายเหตุข้อมูล (ตรวจสอบจากข้อมูลจริงแล้ว):
- ฟิลด์ ASMSTCM.SERIALNO ในระบบนี้ถูกใช้เก็บ "ยี่ห้อ" (เช่น Fujitsu, LENOVO, PANASONIC)
  ไม่ใช่หมายเลขซีเรียลจริงๆ ส่วน ASMSTCM.MANUFACTURER ไม่มีการใช้งาน (ว่างทุกแถว)
- ไม่พบฟิลด์ "รุ่น" (model) แยกต่างหากในฐานข้อมูล จึงเว้นว่างให้กรอกด้วยมือ
- "สถานที่ตั้ง", "สภาพการใช้งาน (ได้/ไม่ได้)", "สถิติการใช้" ไม่มีในฐานข้อมูลเช่นกัน
  เป็นช่องที่ผู้ตรวจกรอกด้วยมือระหว่างเดินตรวจนับจริง (ตามแบบฟอร์มต้นฉบับ)
"""
import io
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from db_connection import run_query

# Business rule ที่หน่วยงานแจ้ง - ตอนนี้เป็นตัวเลือกเปิด/ปิดได้ ไม่บังคับเสมอไป
DEPREGROUP = "1"    # 1 = ขึ้นบัญชีสินทรัพย์
DEPREMETHOD = "1"   # 1 = ค่าเสื่อมราคาแบบเส้นตรง

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]

THAI_FONT = "TH Sarabun New"
_THIN = Side(style="thin", color="000000")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
_GROUP_FILL = PatternFill("solid", fgColor="F2F2F2")

# (label, width) - สภาพการใช้งาน มี 2 ช่องย่อย (ได้ / ไม่ได้) นับรวมเป็น 2 คอลัมน์
_COLUMNS = [
    "ลำดับ",
    "หมายเลขครุภัณฑ์",
    "ชื่อไทย",
    "วันที่ได้มา",
    "สถานที่ตั้ง",
    "ยี่ห้อ",
    "รุ่น",
    "ราคา",
    "ได้",       # สภาพการใช้งาน (ซ้าย)
    "ไม่ได้",     # สภาพการใช้งาน (ขวา)
    "สถิติการใช้",
]
_WIDTHS = [6, 18, 32, 14, 18, 14, 14, 12, 7, 7, 14]
_COND_START_COL = 9   # คอลัมน์ "ได้"
_COND_END_COL = 10    # คอลัมน์ "ไม่ได้"


def to_thai_date(value, with_time: bool = False) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        ts = pd.Timestamp(value)
    else:
        if pd.isna(value):
            return ""
        ts = pd.Timestamp(value)
    text = f"{ts.day} {THAI_MONTHS[ts.month]} {ts.year + 543}"
    if with_time:
        text += f" {ts.hour:02d}:{ts.minute:02d}"
    return text


# ไม่เอาฝ่าย/แผนก/งานที่ชื่อมีคำว่า "เก่า" หรือ "ยกเลิก" มาแสดงให้เลือก (เลิกใช้งานแล้ว)
_EXCLUDE_NAME_SQL = "(NOT ({col} LIKE N'%เก่า%' OR {col} LIKE N'%ยกเลิก%'))"


def list_divisions() -> pd.DataFrame:
    cond = _EXCLUDE_NAME_SQL.format(col="DivisionName")
    return run_query(
        f"""SELECT Division, DivisionName FROM Division
        WHERE (DivisionName IS NULL OR {cond})
          AND Division LIKE '[0-9][0-9][0-9]'
        ORDER BY Division"""
    )


def list_depts(division: str) -> pd.DataFrame:
    cond = _EXCLUDE_NAME_SQL.format(col="DeptName")
    return run_query(
        f"SELECT Dept, DeptName FROM Dept WHERE Division = ? AND (DeptName IS NULL OR {cond}) ORDER BY Dept",
        params=(division,),
    )


def list_sections(division: str, dept: str) -> pd.DataFrame:
    cond = _EXCLUDE_NAME_SQL.format(col="SectionName")
    return run_query(
        f"SELECT Section, SectionName FROM Section WHERE Division = ? AND Dept = ? AND (SectionName IS NULL OR {cond}) ORDER BY Section",
        params=(division, dept),
    )


def fetch_asset_data(
    as_of_date,
    division=None,
    dept=None,
    section=None,
    apply_depre_filter: bool = False,
) -> pd.DataFrame:
    conditions = [
        # a genuine disposal needs BOTH fields; if either is missing (e.g. a transfer
        # cleared DISPOSCODE but left a stale DISPOSDATETIME behind) treat as not disposed.
        # Also: a later transfer record (ASMSTTF) after the disposal date means the asset
        # was brought back into use even though DISPOSCODE/DISPOSDATETIME were never cleared
        "(cm.DISPOSCODE IS NULL OR cm.DISPOSDATETIME IS NULL OR cm.DISPOSDATETIME > ?"
        " OR EXISTS (SELECT 1 FROM ASMSTTF tf WHERE tf.ASSETCODE = cm.ASSETCODE"
        " AND tf.MAKEDATETIME > cm.DISPOSDATETIME))",
        "cm.ACQDATETIME <= ?",
    ]
    params = [as_of_date, as_of_date]

    join_dep = "LEFT JOIN"
    if apply_depre_filter:
        join_dep = "JOIN"
        conditions.append("dep.DEPREGROUP = ?")
        conditions.append("dep.DEPREMETHOD = ?")
        params.extend([DEPREGROUP, DEPREMETHOD])

    if division:
        conditions.append("cm.LOCATEDIVISION = ?")
        params.append(division)
    if dept:
        conditions.append("cm.LOCATEDEPT = ?")
        params.append(dept)
    if section:
        conditions.append("cm.LOCATESECTION = ?")
        params.append(section)

    where_clause = " AND ".join(conditions)

    sql = f"""
    SELECT
        cm.ASSETCODE,
        cm.SUFFIX,
        dbo.GetSSBName(ISNULL(cm.THAINAME, cm.ENGLISHNAME)) AS AssetName,
        cm.SERIALNO AS Brand,
        cm.ACQDATETIME,
        cm.PRICE,
        cm.QTY,
        cm.LOCATEDIVISION,
        cm.LOCATEDEPT,
        cm.LOCATESECTION,
        div.DivisionName,
        dept.DeptName,
        sec.SectionName
    FROM ASMSTCM cm
    {join_dep} ASMSTDEP dep ON dep.ASSETCODE = cm.ASSETCODE AND dep.SUFFIX = cm.SUFFIX
    LEFT JOIN Division div ON cm.LOCATEDIVISION = div.Division
    LEFT JOIN Dept dept ON cm.LOCATEDIVISION = dept.Division AND cm.LOCATEDEPT = dept.Dept
    LEFT JOIN Section sec ON cm.LOCATEDIVISION = sec.Division AND cm.LOCATEDEPT = sec.Dept AND cm.LOCATESECTION = sec.Section
    WHERE {where_clause}
    ORDER BY cm.LOCATEDIVISION, cm.LOCATEDEPT, cm.LOCATESECTION, cm.ASSETCODE
    """

    df = run_query(sql, params=tuple(params))
    df["AcqDateThai"] = df["ACQDATETIME"].apply(to_thai_date)
    df["DivisionLabel"] = df["DivisionName"].fillna("(ไม่ระบุฝ่าย)")
    df["DeptSectionLabel"] = (
        df["DeptName"].fillna("") + " " + df["SectionName"].fillna("")
    ).str.strip()
    df["GroupKey"] = list(zip(df["LOCATEDIVISION"], df["LOCATEDEPT"], df["LOCATESECTION"]))
    return df


def _write_group_header(ws, row, n_cols, division_label, dept_section_label, print_dt):
    third = max(n_cols // 3, 1)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=third)
    c1 = ws.cell(row, 1, division_label)
    c1.font = Font(name=THAI_FONT, size=13, bold=True)
    c1.alignment = Alignment(horizontal="left")

    ws.merge_cells(start_row=row, start_column=third + 1, end_row=row, end_column=n_cols - third)
    c2 = ws.cell(row, third + 1, dept_section_label)
    c2.font = Font(name=THAI_FONT, size=13, bold=True)
    c2.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=row, start_column=n_cols - third + 1, end_row=row, end_column=n_cols)
    c3 = ws.cell(row, n_cols - third + 1, f"วันที่พิมพ์ {to_thai_date(print_dt, with_time=True)}")
    c3.font = Font(name=THAI_FONT, size=11)
    c3.alignment = Alignment(horizontal="right")
    return row + 1


def _write_column_headers(ws, row):
    r1, r2 = row, row + 1
    for col_idx, label in enumerate(_COLUMNS, start=1):
        if col_idx in (_COND_START_COL, _COND_END_COL):
            continue
        ws.merge_cells(start_row=r1, start_column=col_idx, end_row=r2, end_column=col_idx)
        cell = ws.cell(r1, col_idx, label)
        cell.font = Font(name=THAI_FONT, size=13, bold=True)
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDER
        ws.cell(r2, col_idx).border = _BORDER
        ws.cell(r2, col_idx).fill = _HEADER_FILL

    ws.merge_cells(start_row=r1, start_column=_COND_START_COL, end_row=r1, end_column=_COND_END_COL)
    cond_cell = ws.cell(r1, _COND_START_COL, "สภาพการใช้งาน")
    cond_cell.font = Font(name=THAI_FONT, size=13, bold=True)
    cond_cell.fill = _HEADER_FILL
    cond_cell.alignment = Alignment(horizontal="center", vertical="center")
    cond_cell.border = _BORDER
    ws.cell(r1, _COND_END_COL).border = _BORDER
    ws.cell(r1, _COND_END_COL).fill = _HEADER_FILL

    for col_idx in (_COND_START_COL, _COND_END_COL):
        cell = ws.cell(r2, col_idx, _COLUMNS[col_idx - 1])
        cell.font = Font(name=THAI_FONT, size=11, bold=True)
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _BORDER

    return r2 + 1


def build_excel_bytes(df: pd.DataFrame, fiscal_year: str, fiscal_period: str, as_of_date) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "รายงานตรวจครุภัณฑ์"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddFooter.right.text = "หน้า &P / &N"
    ws.oddFooter.right.size = 10

    n_cols = len(_COLUMNS)
    for col_idx, width in enumerate(_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    print_dt = datetime.now()

    # ==== ชื่อรายงาน ====
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(1, 1, "บัญชีรายการครุภัณฑ์")
    title_cell.font = Font(name=THAI_FONT, size=20, bold=True)
    title_cell.alignment = Alignment(horizontal="center")

    subtitle_bits = []
    if fiscal_year:
        subtitle_bits.append(f"ปี {fiscal_year}")
    if fiscal_period:
        subtitle_bits.append(f"งวด {fiscal_period}")
    subtitle_bits.append(f"ข้อมูล ณ วันที่ {to_thai_date(as_of_date)}")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    sub_cell = ws.cell(2, 1, "  ".join(subtitle_bits))
    sub_cell.font = Font(name=THAI_FONT, size=13)
    sub_cell.alignment = Alignment(horizontal="center")

    row = 3
    grand_total_qty = 0
    grand_total_price = 0.0

    if df.empty:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        ws.cell(row, 1, "ไม่พบรายการครุภัณฑ์ตามเงื่อนไขที่เลือก")
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.getvalue()

    for group_key, gdf in df.groupby("GroupKey", sort=False):
        division_label = gdf["DivisionLabel"].iloc[0]
        dept_section_label = gdf["DeptSectionLabel"].iloc[0]

        row = _write_group_header(ws, row, n_cols, division_label, dept_section_label, print_dt)
        row = _write_column_headers(ws, row)

        seq = 0
        group_total = 0.0
        for _, r in gdf.iterrows():
            seq += 1
            row_values = {
                1: seq,
                2: r["ASSETCODE"],
                3: r["AssetName"],
                4: r["AcqDateThai"],
                5: "",  # สถานที่ตั้ง - กรอกด้วยมือระหว่างตรวจนับ
                6: r["Brand"] or "",
                7: "",  # รุ่น - ไม่มีในฐานข้อมูล กรอกด้วยมือ
                8: r["PRICE"] if pd.notna(r["PRICE"]) and r["PRICE"] else None,
                9: "",   # ได้ - กรอกด้วยมือ
                10: "",  # ไม่ได้ - กรอกด้วยมือ
                11: "",  # สถิติการใช้ - กรอกด้วยมือ
            }
            for col_idx, value in row_values.items():
                cell = ws.cell(row, col_idx, value)
                cell.font = Font(name=THAI_FONT, size=13)
                cell.border = _BORDER
                if col_idx in (1,):
                    cell.alignment = Alignment(horizontal="center")
                elif col_idx == 4:
                    cell.alignment = Alignment(horizontal="center")
                elif col_idx == 8:
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right")
                elif col_idx in (9, 10):
                    cell.alignment = Alignment(horizontal="center")
            if pd.notna(r["PRICE"]) and r["PRICE"]:
                group_total += float(r["PRICE"])
            row += 1

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        tcell = ws.cell(row, 1, f"รวม {division_label} {dept_section_label} ({seq} รายการ)")
        tcell.font = Font(name=THAI_FONT, size=13, bold=True)
        tcell.alignment = Alignment(horizontal="right")
        pcell = ws.cell(row, 8, group_total)
        pcell.font = Font(name=THAI_FONT, size=13, bold=True)
        pcell.number_format = "#,##0.00"
        pcell.alignment = Alignment(horizontal="right")
        for c in range(1, n_cols + 1):
            ws.cell(row, c).border = _BORDER
            ws.cell(row, c).fill = _GROUP_FILL
        row += 2

        grand_total_qty += seq
        grand_total_price += group_total

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    gcell = ws.cell(row, 1, f"รวมทั้งสิ้น {grand_total_qty} รายการ มูลค่า")
    gcell.font = Font(name=THAI_FONT, size=15, bold=True)
    gcell.alignment = Alignment(horizontal="right")
    gpcell = ws.cell(row, 8, grand_total_price)
    gpcell.font = Font(name=THAI_FONT, size=15, bold=True)
    gpcell.number_format = "#,##0.00"
    gpcell.alignment = Alignment(horizontal="right")
    ws.cell(row, 9, "บาท").font = Font(name=THAI_FONT, size=13)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
