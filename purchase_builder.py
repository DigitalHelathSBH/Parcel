"""
รายงานซื้อประจำเดือน (อย่างง่าย) - ข้อมูลเตรียมก่อนกรอกแบบ สขร.1
ดึงจาก APINV/APINVITEM (ใบสั่งซื้อ/ใบแจ้งหนี้) ตามช่วงต้นเดือนถึงสิ้นเดือนที่เลือก (PODATETIME)

หมายเหตุสำคัญ: ฐานข้อมูลนี้ไม่มีข้อมูล "วิธีซื้อหรือจ้าง" / "รายชื่อผู้เสนอราคา" /
"เหตุผลที่คัดเลือกโดยสรุป" ตามแบบ สขร.1 มาตรฐานราชการ จึงเว้น 3 คอลัมน์นี้ไว้ว่าง
ให้กรอกด้วยมือก่อนเผยแพร่จริง ไฟล์นี้สร้างได้แค่รายการซื้อพื้นฐาน (เลขที่ PO, วันที่,
ผู้ขาย, รายการ, จำนวนเงิน, หมวดงบประมาณ) เท่านั้น
"""
import calendar
import io
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from db_connection import run_query
from depreciation_builder import THAI_FONT, _BORDER, _HEADER_FILL, to_thai_date

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]


def fetch_purchase_data(calendar_year: int, calendar_month: int) -> pd.DataFrame:
    start_date = date(calendar_year, calendar_month, 1)
    if calendar_month == 12:
        end_date = date(calendar_year + 1, 1, 1)
    else:
        end_date = date(calendar_year, calendar_month + 1, 1)

    sql = """
    SELECT
        inv.PONO,
        inv.PODATETIME,
        inv.APCODE,
        dbo.GetSSBName(ap.THAINAME) AS VendorName,
        inv.PURCHASEBUDGETCATEGORY,
        SUM(item.AMT) AS TotalAmt,
        MAX(item.NAME) AS ItemName,
        MAX(item.DESCRIPTION) AS ItemDescription
    FROM APINV inv
    JOIN APINVITEM item ON item.INVOICENO = inv.INVOICENO AND item.APCODE = inv.APCODE
    LEFT JOIN APMASTER ap ON ap.APCODE = inv.APCODE
    WHERE inv.PODATETIME >= ? AND inv.PODATETIME < ?
      AND inv.VOIDDATETIME IS NULL
    GROUP BY inv.PONO, inv.PODATETIME, inv.APCODE, ap.THAINAME, inv.PURCHASEBUDGETCATEGORY
    ORDER BY inv.PODATETIME, inv.PONO
    """
    df = run_query(sql, params=(start_date, end_date))
    if df.empty:
        return df

    df["PODateThai"] = df["PODATETIME"].apply(to_thai_date)
    df["VendorLabel"] = df["VendorName"].fillna(df["APCODE"])
    df["ItemLabel"] = df["ItemName"].fillna(df["ItemDescription"]).fillna("")
    return df


_COLUMNS = [
    "ลำดับ", "เลขที่ใบสั่งซื้อ/ใบแจ้งหนี้", "วันที่", "ผู้ขาย", "รายการที่จัดซื้อ/จัดจ้าง",
    "จำนวนเงิน", "หมวดงบประมาณ",
    "วิธีซื้อหรือจ้าง", "ผู้เสนอราคาและราคาที่เสนอ", "เหตุผลที่คัดเลือกโดยสรุป",
]
_WIDTHS = [6, 20, 12, 26, 28, 14, 14, 16, 24, 24]


def build_excel_bytes(df: pd.DataFrame, calendar_year: int, calendar_month: int) -> bytes:
    month_label = THAI_MONTHS[calendar_month]
    fiscal_be = calendar_year + 543

    wb = Workbook()
    ws = wb.active
    ws.title = "รายงานซื้อประจำเดือน"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    n_cols = len(_COLUMNS)
    for i, w in enumerate(_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(1, 1, "สรุปผลการจัดซื้อจัดจ้างในรอบเดือน (ข้อมูลเบื้องต้นจากระบบ - ยังไม่ใช่แบบ สขร.1 ที่สมบูรณ์)")
    c.font = Font(name=THAI_FONT, size=16, bold=True)
    c.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    c2 = ws.cell(2, 1, f"ประจำเดือน{month_label} {fiscal_be}")
    c2.font = Font(name=THAI_FONT, size=13)
    c2.alignment = Alignment(horizontal="center")

    row = 4
    if df.empty:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        ws.cell(row, 1, "ไม่พบข้อมูลตามเงื่อนไขที่เลือก")
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.getvalue()

    for col_idx, label in enumerate(_COLUMNS, start=1):
        cell = ws.cell(row, col_idx, label)
        cell.font = Font(name=THAI_FONT, size=12, bold=True)
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = _BORDER
    row += 1

    grand_total = 0.0
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        values = [
            i, r["PONO"], r["PODateThai"], r["VendorLabel"], r["ItemLabel"],
            r["TotalAmt"], r["PURCHASEBUDGETCATEGORY"] or "",
            "", "", "",
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row, col_idx, value)
            cell.font = Font(name=THAI_FONT, size=12)
            cell.border = _BORDER
            if col_idx == 6:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx in (1, 3):
                cell.alignment = Alignment(horizontal="center")
        grand_total += float(r["TotalAmt"] or 0.0)
        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    gcell = ws.cell(row, 1, f"รวมทั้งสิ้น ({len(df)} รายการ)")
    gcell.font = Font(name=THAI_FONT, size=13, bold=True)
    gcell.alignment = Alignment(horizontal="right")
    gc = ws.cell(row, 6, grand_total)
    gc.font = Font(name=THAI_FONT, size=13, bold=True)
    gc.number_format = "#,##0.00"
    gc.alignment = Alignment(horizontal="right")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
