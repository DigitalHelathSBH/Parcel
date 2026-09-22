"""
รายงานครุภัณฑ์ตัดจำหน่าย รายละเอียด ตามแหล่งเงิน
เลย์เอาต์เดียวกับรายงานค่าเสื่อมราคารายละเอียดตามแหล่งเงิน (depreciation_builder.build_detail_excel_bytes)
แต่กรองเฉพาะรายการที่ตัดจำหน่ายแล้วในช่วงเดือน/ปีที่เลือก แทนคอลัมน์ค่าเสื่อมด้วยวันที่/เหตุผลตัดจำหน่าย
"""
import calendar
import io
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from db_connection import run_query
from depreciation_builder import (
    THAI_FONT,
    THAI_MONTHS,
    _BORDER,
    _GROUP_FILL,
    _HEADER_FILL,
    _SUBGROUP_FILL,
    to_thai_date,
)


def fetch_disposal_data(calendar_year: int, calendar_month: int, division=None, dept=None, section=None) -> pd.DataFrame:
    start_date = date(calendar_year, calendar_month, 1)
    if calendar_month == 12:
        end_date = date(calendar_year + 1, 1, 1)
    else:
        end_date = date(calendar_year, calendar_month + 1, 1)

    conditions = [
        "cm.DISPOSDATETIME IS NOT NULL",
        "cm.DISPOSCODE IS NOT NULL",
        "cm.DISPOSDATETIME >= ?",
        "cm.DISPOSDATETIME < ?",
    ]
    params = [start_date, end_date]

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
        dbo.GetSSBName(ISNULL(cm.THAINAME, m.THAINAME)) AS AssetName,
        COALESCE(cm.ACQDATETIME, m.FIRSTDATETIME) AS AcqDate,
        cm.QTY,
        cm.PRICE,
        cm.DISPOSDATETIME,
        cm.DISPOSCODE,
        dbo.GetSSBName(dc.THAINAME) AS DisposeReasonName,
        cm.LOCATEDIVISION,
        div.DivisionName,
        cm.ARTICLEGROUP,
        dbo.GetSSBName(ag.THAINAME) AS ArticleGroupName,
        cm.PURCHASEBUDGETCODE,
        dbo.GetSSBName(bg.THAINAME) AS BudgetName
    FROM ASMSTCM cm
    JOIN ASMST m ON m.ASSETCODE = cm.ASSETCODE
    LEFT JOIN Division div ON cm.LOCATEDIVISION = div.Division
    LEFT JOIN SYSCONFIG ag ON ag.CODE = cm.ARTICLEGROUP AND ag.CTRLCODE = 100012
    LEFT JOIN SYSCONFIG bg ON bg.CODE = cm.PURCHASEBUDGETCODE AND bg.CTRLCODE = 120010
    LEFT JOIN SYSCONFIG dc ON dc.CODE = cm.DISPOSCODE AND dc.CTRLCODE = 100011
    WHERE {where_clause}
    ORDER BY cm.PURCHASEBUDGETCODE, cm.ARTICLEGROUP, cm.DISPOSDATETIME
    """
    df = run_query(sql, params=tuple(params))
    if df.empty:
        return df

    df["TotalValue"] = df["QTY"].fillna(0.0) * df["PRICE"].fillna(0.0)
    df["AcqDateThai"] = df["AcqDate"].apply(to_thai_date)
    df["DisposeDateThai"] = df["DISPOSDATETIME"].apply(to_thai_date)
    df["DisposeReasonLabel"] = df["DisposeReasonName"].fillna(df["DISPOSCODE"])
    df["BudgetLabel"] = df["BudgetName"].fillna("(ไม่ระบุแหล่งเงิน)")
    df["ArticleGroupLabel"] = (
        "(" + df["ARTICLEGROUP"].fillna("") + ") " + df["ArticleGroupName"].fillna("")
    )
    df["DivisionLabel"] = df["DivisionName"].fillna("(ไม่ระบุฝ่าย)")
    return df


_COLUMNS = ["รหัส", "รายการ", "ฝ่าย", "วันที่ได้มา", "มูลค่า", "วันที่ตัดจำหน่าย", "เหตุผลที่ตัดจำหน่าย"]
_WIDTHS = [16, 36, 24, 14, 14, 16, 26]


def build_excel_bytes(df: pd.DataFrame, calendar_year: int, calendar_month: int) -> bytes:
    fiscal_be = calendar_year + 543
    month_label = THAI_MONTHS[calendar_month]

    wb = Workbook()
    ws = wb.active
    ws.title = "ตัดจำหน่าย"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    n_cols = len(_COLUMNS)
    for i, w in enumerate(_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(1, 1, "รายงานครุภัณฑ์ตัดจำหน่าย รายละเอียด ตามแหล่งเงิน")
    c.font = Font(name=THAI_FONT, size=18, bold=True)
    c.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    c2 = ws.cell(2, 1, f"ตัดจำหน่ายในเดือน{month_label} {fiscal_be}")
    c2.font = Font(name=THAI_FONT, size=13)
    c2.alignment = Alignment(horizontal="center")
    row = 4

    if df.empty:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        ws.cell(row, 1, "ไม่พบรายการตัดจำหน่ายตามเงื่อนไขที่เลือก")
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.getvalue()

    grand_total_value = 0.0
    grand_count = 0

    for budget_label, bdf in df.groupby("BudgetLabel", sort=False):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        gc = ws.cell(row, 1, budget_label)
        gc.font = Font(name=THAI_FONT, size=14, bold=True)
        gc.fill = _GROUP_FILL
        row += 1

        for art_label, adf in bdf.groupby("ArticleGroupLabel", sort=False):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
            ac = ws.cell(row, 1, art_label)
            ac.font = Font(name=THAI_FONT, size=13, bold=True)
            ac.fill = _SUBGROUP_FILL
            row += 1

            for col_idx, label in enumerate(_COLUMNS, start=1):
                cell = ws.cell(row, col_idx, label)
                cell.font = Font(name=THAI_FONT, size=12, bold=True)
                cell.fill = _HEADER_FILL
                cell.alignment = Alignment(horizontal="center")
                cell.border = _BORDER
            row += 1

            sub_total_value = 0.0
            for _, r in adf.iterrows():
                values = [
                    r["ASSETCODE"], r["AssetName"], r["DivisionLabel"], r["AcqDateThai"],
                    r["TotalValue"], r["DisposeDateThai"], r["DisposeReasonLabel"],
                ]
                for col_idx, value in enumerate(values, start=1):
                    cell = ws.cell(row, col_idx, value)
                    cell.font = Font(name=THAI_FONT, size=12)
                    cell.border = _BORDER
                    if col_idx == 5:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right")
                    elif col_idx in (4, 6):
                        cell.alignment = Alignment(horizontal="center")
                sub_total_value += float(r["TotalValue"])
                row += 1

            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            tc = ws.cell(row, 1, f"รวม {art_label} ({len(adf)} รายการ)")
            tc.font = Font(name=THAI_FONT, size=12, bold=True)
            tc.alignment = Alignment(horizontal="right")
            c = ws.cell(row, 5, sub_total_value)
            c.font = Font(name=THAI_FONT, size=12, bold=True)
            c.number_format = "#,##0.00"
            c.alignment = Alignment(horizontal="right")
            row += 2

            grand_total_value += sub_total_value
            grand_count += len(adf)

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    gcell = ws.cell(row, 1, f"รวมทั้งสิ้น ({grand_count} รายการ)")
    gcell.font = Font(name=THAI_FONT, size=14, bold=True)
    gcell.alignment = Alignment(horizontal="right")
    gc2 = ws.cell(row, 5, grand_total_value)
    gc2.font = Font(name=THAI_FONT, size=14, bold=True)
    gc2.number_format = "#,##0.00"
    gc2.alignment = Alignment(horizontal="right")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
