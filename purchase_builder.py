"""
รายงานซื้อประจำเดือน - ใช้เตรียมข้อมูลก่อนกรอกแบบ สขร.1
ดึงจาก SKPO/SKPODTL (ใบสั่งซื้อจริงของระบบพัสดุ) ตามช่วงวันที่ออกใบสั่งซื้อ (ISSUEDATETIME)
ในเดือนที่เลือก จัดกลุ่มตามประเภทค่าใช้จ่าย (จากชื่อรหัสบัญชี StockActView) แล้วตามด้วย
ใบสั่งซื้อแต่ละใบ และตารางรายการย่อยแยกรายบรรทัดภายใต้ใบสั่งซื้อนั้น

หมายเหตุ: "วิธีซื้อหรือจ้าง" มาจาก SKPO.PURCHASETYPECODE (SYSCONFIG CTRLCODE=100010)
ซึ่งเป็นข้อมูลจริงที่ระบบเก็บไว้ (ไม่ใช่ช่องว่างให้กรอกเองเหมือนก่อนหน้านี้)
"""
import io
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
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


_DRUG_MAINCATEGORIES = {
    "ยาในบัญชียาหลัก (ED)", "ยานอกบัญชียาหลัก (NE)",
    "ยากรณีพิเศษ ในบัญชียาหลักแห่งชาติ", "ยากรณีพิเศษ นอกบัญชียาหลักแห่งชาติ", "ยาอื่นๆ",
}
_MATERIAL_MAINCATEGORIES = {
    "วัสดุ", "เวชภัณฑ์ที่มิใช่ยา", "วัสดุการแพทย์ (MS)", "น้ำยาห้องปฏิบัติการ (LAB)",
    "สารเคมีห้องปฏิบัติการ", "วัสดุการแพทย์ทันตกรรม",
}


def _classify_category(stockactname: str, maincategoryname: str) -> str:
    """
    แยกประเภทค่าใช้จ่าย - ก่อนอื่นดูชื่อรหัสบัญชี (StockActView.stockactname) ซึ่งบอกประเภทชัดเจนที่สุด
    แต่บางรหัสเป็นรหัสกลางๆ ไม่บอกประเภท (เช่น "ซื้อรวมภาษี(วิธีเฉพาะเจาะจง)") จึง fallback ไปดู
    หมวดของตัววัสดุเอง (STOCK_MASTER.MAINCATEGORY ผ่าน StockMainCategory) แทน
    """
    name = stockactname or ""
    if "ต่ำกว่าเกณฑ์" in name:
        return "ครุภัณฑ์ต่ำกว่าเกณฑ์"
    if "ซ่อมแซม" in name:
        return "ค่าซ่อมแซม"
    if "จ้างเหมา" in name:
        return "ค่าจ้างเหมาบริการ"
    if "ครุภัณฑ์" in name:
        if "คอมพิวเตอร์" in name:
            return "ครุภัณฑ์คอมพิวเตอร์"
        if "การแพทย์" in name or "วิทยาศาสตร์" in name:
            return "ครุภัณฑ์วิทยาศาสตร์และการแพทย์"
        if "สำนักงาน" in name:
            return "ครุภัณฑ์สำนักงาน"
        return "ครุภัณฑ์อื่นๆ"
    if "วัสดุ" in name:
        return "วัสดุ"
    if "สาธารณูปโภค" in name:
        return "ค่าสาธารณูปโภค"
    if "เบ็ดเตล็ด" in name:
        return "ค่าใช้จ่ายเบ็ดเตล็ด"

    mc = maincategoryname or ""
    if mc in _DRUG_MAINCATEGORIES:
        return "ยา"
    if mc in _MATERIAL_MAINCATEGORIES:
        return "วัสดุ"
    if mc == "ครุภัณฑ์":
        return "ครุภัณฑ์อื่นๆ"
    if mc == "งานจ้าง":
        return "ค่าจ้างเหมาบริการ"
    if mc == "ค่าสาธารณูปโภค":
        return "ค่าสาธารณูปโภค"
    return "(ไม่ระบุประเภท)"


def fetch_purchase_data(calendar_year: int, calendar_month: int) -> pd.DataFrame:
    start_date = date(calendar_year, calendar_month, 1)
    if calendar_month == 12:
        end_date = date(calendar_year + 1, 1, 1)
    else:
        end_date = date(calendar_year, calendar_month + 1, 1)

    sql = """
    SELECT
        po.PONO,
        po.ISSUEDATETIME,
        po.SUPPLIERCODE,
        dbo.GetSSBName(ap.THAINAME) AS VendorName,
        po.PURCHASETYPECODE,
        dbo.GetSSBName(pt.THAINAME) AS PurchaseTypeName,
        po.PURCHASEBUDGETCODE,
        dbo.GetSSBName(bg.THAINAME) AS BudgetName,
        dt.STOCKCODE,
        dbo.GetSSBName(sm.THAINAME) AS ItemName,
        dt.UNITCODE,
        uc.UnitName,
        dt.STOCKACTCODE,
        sav.stockactname AS StockActName,
        smc.MainCategoryName,
        dt.REQUESTQTY,
        dt.LOTPRICE,
        dt.AMT,
        dt.GOODSAMTAFTERITEMDISCOUNT,
        dt.ALLOCATEDVATAMT
    FROM SKPO po
    JOIN SKPODTL dt ON dt.PONO = po.PONO
    LEFT JOIN APMASTER ap ON ap.APCODE = po.SUPPLIERCODE
    LEFT JOIN SYSCONFIG pt ON pt.CODE = po.PURCHASETYPECODE AND pt.CTRLCODE = 100010
    LEFT JOIN SYSCONFIG bg ON bg.CODE = po.PURCHASEBUDGETCODE AND bg.CTRLCODE = 120010
    LEFT JOIN STOCK_MASTER sm ON sm.STOCKCODE = dt.STOCKCODE
    LEFT JOIN StockActView sav ON sav.stockactcode = dt.STOCKACTCODE
    LEFT JOIN StockMainCategory smc ON smc.MainCategory = sm.MAINCATEGORY
    LEFT JOIN UnitCode uc ON uc.UnitCode = dt.UNITCODE
    WHERE po.ISSUEDATETIME >= ? AND po.ISSUEDATETIME < ?
      AND po.CXLDATETIME IS NULL
      AND po.STORE = '1'  -- คลังสินค้า "พัสดุ 1" เท่านั้น (ตามที่หน่วยงานระบุ)
      -- ไม่กรอง POSTATUS (เทียบเท่า "All" ในหน้าจอ Check List P/O เดิม)
    ORDER BY po.ISSUEDATETIME, po.PONO, dt.SUFFIX
    """
    df = run_query(sql, params=(start_date, end_date))
    if df.empty:
        return df

    df["PODateThai"] = df["ISSUEDATETIME"].apply(to_thai_date)
    df["VendorLabel"] = df["VendorName"].fillna(df["SUPPLIERCODE"])
    df["PurchaseTypeLabel"] = df["PurchaseTypeName"].fillna("(ไม่ระบุวิธี)")
    # "หมวดเงิน" ในรายงานต้นฉบับ (SKPOOUT-รายละเอียดใบบันทึกข้อความ) คือชื่อรหัสบัญชี
    # (STOCKACTCODE ผ่าน StockActView) ไม่ใช่แหล่งเงินของทั้งใบสั่งซื้อ (PURCHASEBUDGETCODE)
    df["ExpenseAccountLabel"] = df["StockActName"].fillna(df["STOCKACTCODE"])
    df["UnitLabel"] = df["UnitName"].fillna(df["UNITCODE"])
    df["ItemLabel"] = df["ItemName"].fillna(df["STOCKCODE"])
    df["NetAmt"] = df["GOODSAMTAFTERITEMDISCOUNT"].fillna(0.0) + df["ALLOCATEDVATAMT"].fillna(0.0)
    df["PurchaseCategory"] = df.apply(
        lambda r: _classify_category(r["StockActName"], r["MainCategoryName"]), axis=1
    )
    return df


_CATEGORY_FILL = PatternFill("solid", fgColor="2563EB")

_DETAIL_COLUMNS = [
    "ที่", "รหัสวัสดุ", "ชื่อวัสดุ", "หน่วย", "หมวดเงิน", "จำนวน", "ราคา",
    "จำนวนเงิน", "มูลค่าสินค้า", "ภาษีมูลค่าเพิ่ม", "เงินสุทธิ",
]
_WIDTHS = [5, 12, 30, 8, 16, 8, 12, 12, 12, 12, 12]


def build_excel_bytes(df: pd.DataFrame, calendar_year: int, calendar_month: int) -> bytes:
    month_label = THAI_MONTHS[calendar_month]
    fiscal_be = calendar_year + 543
    n_cols = len(_DETAIL_COLUMNS)

    wb = Workbook()
    ws = wb.active
    ws.title = "รายงานซื้อประจำเดือน"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    for i, w in enumerate(_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(1, 1, "รายงานสรุปผลการจัดซื้อจัดจ้างในรอบเดือน")
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

    grand_count = 0
    grand_total = 0.0

    for category, cdf in df.groupby("PurchaseCategory", sort=False):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        catcell = ws.cell(row, 1, category)
        catcell.font = Font(name=THAI_FONT, size=16, bold=True, color="FFFFFF")
        catcell.fill = _CATEGORY_FILL
        catcell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 26
        row += 1

        cat_total = 0.0
        cat_po_count = 0

        po_index = 0
        for (pono, issue_date, vendor, ptype), pdf in cdf.groupby(
            ["PONO", "PODateThai", "VendorLabel", "PurchaseTypeLabel"], sort=False
        ):
            po_index += 1
            cat_po_count += 1
            po_total = float(pdf["AMT"].fillna(0.0).sum())
            cat_total += po_total

            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
            header_text = (
                f"{po_index}. เลขที่ใบสั่งซื้อ/ใบแจ้งหนี้ {pono}    วันที่ {issue_date}    "
                f"ผู้ขาย {vendor}    วิธีซื้อหรือจ้าง {ptype}    จำนวนเงิน {po_total:,.2f} บาท"
            )
            hc = ws.cell(row, 1, header_text)
            hc.font = Font(name=THAI_FONT, size=13, bold=True)
            hc.fill = _SUBGROUP_FILL
            row += 1

            for col_idx, label in enumerate(_DETAIL_COLUMNS, start=1):
                cell = ws.cell(row, col_idx, label)
                cell.font = Font(name=THAI_FONT, size=11, bold=True)
                cell.fill = _HEADER_FILL
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
                cell.border = _BORDER
            row += 1

            for i, (_, r) in enumerate(pdf.iterrows(), start=1):
                values = [
                    i, r["STOCKCODE"], r["ItemLabel"], r["UnitLabel"], r["ExpenseAccountLabel"],
                    r["REQUESTQTY"], r["LOTPRICE"], r["AMT"],
                    r["GOODSAMTAFTERITEMDISCOUNT"], r["ALLOCATEDVATAMT"], r["NetAmt"],
                ]
                for col_idx, value in enumerate(values, start=1):
                    cell = ws.cell(row, col_idx, value)
                    cell.font = Font(name=THAI_FONT, size=11)
                    cell.border = _BORDER
                    if col_idx >= 6:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right")
                    elif col_idx == 1:
                        cell.alignment = Alignment(horizontal="center")
                row += 1

            row += 1

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols - 1)
        subtotal_cell = ws.cell(row, 1, f"รวม {category} ({cat_po_count} ใบสั่งซื้อ)")
        subtotal_cell.font = Font(name=THAI_FONT, size=13, bold=True)
        subtotal_cell.alignment = Alignment(horizontal="right")
        sc = ws.cell(row, n_cols, cat_total)
        sc.font = Font(name=THAI_FONT, size=13, bold=True)
        sc.number_format = "#,##0.00"
        sc.alignment = Alignment(horizontal="right")
        row += 2

        grand_count += cat_po_count
        grand_total += cat_total

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols - 1)
    gcell = ws.cell(row, 1, f"รวมทั้งสิ้น ({grand_count} ใบสั่งซื้อ)")
    gcell.font = Font(name=THAI_FONT, size=14, bold=True)
    gcell.alignment = Alignment(horizontal="right")
    gc2 = ws.cell(row, n_cols, grand_total)
    gc2.font = Font(name=THAI_FONT, size=14, bold=True)
    gc2.number_format = "#,##0.00"
    gc2.alignment = Alignment(horizontal="right")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


_SUMMARY_COLUMNS = ["ที่", "เลขที่ใบสั่งซื้อ/ใบแจ้งหนี้", "วันที่", "ผู้ขาย", "จำนวนรายการ", "เงินรวมสุทธิ"]
_SUMMARY_WIDTHS = [6, 20, 14, 34, 14, 16]


def build_summary_excel_bytes(df: pd.DataFrame, calendar_year: int, calendar_month: int) -> bytes:
    month_label = THAI_MONTHS[calendar_month]
    fiscal_be = calendar_year + 543
    n_cols = len(_SUMMARY_COLUMNS)

    wb = Workbook()
    ws = wb.active
    ws.title = "สรุปรายงานซื้อประจำเดือน"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    for i, w in enumerate(_SUMMARY_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(1, 1, "รายงานสรุปผลการจัดซื้อจัดจ้างในรอบเดือน (สรุปตามใบสั่งซื้อ)")
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

    po_agg = df.groupby(["PONO", "PODateThai", "VendorLabel", "PurchaseCategory"], sort=False).agg(
        ItemCount=("STOCKCODE", "count"),
        NetTotal=("NetAmt", "sum"),
    ).reset_index()

    grand_count = 0
    grand_total = 0.0

    for category, cdf in po_agg.groupby("PurchaseCategory", sort=False):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        catcell = ws.cell(row, 1, category)
        catcell.font = Font(name=THAI_FONT, size=16, bold=True, color="FFFFFF")
        catcell.fill = _CATEGORY_FILL
        catcell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 26
        row += 1

        for col_idx, label in enumerate(_SUMMARY_COLUMNS, start=1):
            cell = ws.cell(row, col_idx, label)
            cell.font = Font(name=THAI_FONT, size=12, bold=True)
            cell.fill = _HEADER_FILL
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = _BORDER
        row += 1

        cat_total = 0.0
        for i, (_, r) in enumerate(cdf.iterrows(), start=1):
            values = [i, r["PONO"], r["PODateThai"], r["VendorLabel"], r["ItemCount"], r["NetTotal"]]
            for col_idx, value in enumerate(values, start=1):
                cell = ws.cell(row, col_idx, value)
                cell.font = Font(name=THAI_FONT, size=12)
                cell.border = _BORDER
                if col_idx == 6:
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right")
                elif col_idx in (1, 5):
                    cell.alignment = Alignment(horizontal="center")
            cat_total += float(r["NetTotal"])
            row += 1

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols - 1)
        subtotal_cell = ws.cell(row, 1, f"รวม {category} ({len(cdf)} ใบสั่งซื้อ)")
        subtotal_cell.font = Font(name=THAI_FONT, size=13, bold=True)
        subtotal_cell.alignment = Alignment(horizontal="right")
        sc = ws.cell(row, n_cols, cat_total)
        sc.font = Font(name=THAI_FONT, size=13, bold=True)
        sc.number_format = "#,##0.00"
        sc.alignment = Alignment(horizontal="right")
        row += 2

        grand_count += len(cdf)
        grand_total += cat_total

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols - 1)
    gcell = ws.cell(row, 1, f"รวมทั้งสิ้น ({grand_count} ใบสั่งซื้อ)")
    gcell.font = Font(name=THAI_FONT, size=14, bold=True)
    gcell.alignment = Alignment(horizontal="right")
    gc2 = ws.cell(row, n_cols, grand_total)
    gc2.font = Font(name=THAI_FONT, size=14, bold=True)
    gc2.number_format = "#,##0.00"
    gc2.alignment = Alignment(horizontal="right")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


_THAI_MONTHS_ABBR = [
    "", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค.",
]


def _to_thai_date_short(value) -> str:
    if value is None or pd.isna(value):
        return ""
    ts = pd.Timestamp(value)
    return f"{ts.day} {_THAI_MONTHS_ABBR[ts.month]} {(ts.year + 543) % 100:02d}"


# คอลัมน์สุดท้าย 2 อัน (วันที่, เลขที่) ใช้หัวข้อร่วมกัน "เลขที่และวันที่ของสัญญาหรือข้อตกลง" (merge)
_SAKOR_COLUMNS = [
    "ลำดับที่", "งานที่จัดซื้อหรือจัดจ้าง", "วงเงินที่จะซื้อหรือจ้าง", "ราคากลาง",
    "วิธีซื้อหรือจ้าง", "รายชื่อผู้เสนอราคา", "ราคาที่เสนอ", "ผู้ได้รับการคัดเลือก",
    "ราคาที่ตกลงซื้อหรือจ้าง", "เหตุผลที่คัดเลือกโดยสรุป", "เลขที่", "วันที่",
]
_SAKOR_WIDTHS = [6, 26, 14, 14, 16, 24, 14, 24, 14, 30, 14, 14]
_SAKOR_CONTRACT_HEADER = "เลขที่และวันที่ของสัญญาหรือข้อตกลง"
_SAKOR_REASON_TEXT = "เป็นผู้มีคุณสมบัติถูกต้องตามเงื่อนไขในการตกลงราคา"
_SAKOR_EBIDDING_THRESHOLD = 500000


def build_sakor_excel_bytes(df: pd.DataFrame, calendar_year: int, calendar_month: int) -> bytes:
    """
    แบบสรุปผลการดำเนินการจัดซื้อจัดจ้างในรอบเดือน (แบบ สขร.1) - รายการเรียงตามวันที่ ไม่แบ่งกลุ่มตามประเภท
    เพราะแบบฟอร์มราชการนี้เป็นรายการต่อเนื่องตามลำดับเวลา ไม่ได้จัดกลุ่มตามหมวดหมู่
    """
    month_label = THAI_MONTHS[calendar_month]
    fiscal_be = calendar_year + 543
    n_cols = len(_SAKOR_COLUMNS)

    wb = Workbook()
    ws = wb.active
    ws.title = "สขร.1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    for i, w in enumerate(_SAKOR_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(1, 1, f"แบบสรุปผลการดำเนินการจัดซื้อจัดจ้างในรอบเดือน {month_label} {fiscal_be}")
    c.font = Font(name=THAI_FONT, size=16, bold=True)
    c.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    c2 = ws.cell(2, 1, "โรงพยาบาลสระบุรี")
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

    for col_idx, label in enumerate(_SAKOR_COLUMNS[:-2], start=1):
        cell = ws.cell(row, col_idx, label)
        cell.font = Font(name=THAI_FONT, size=12, bold=True)
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = _BORDER

    ws.merge_cells(start_row=row, start_column=n_cols - 1, end_row=row, end_column=n_cols)
    contract_header = ws.cell(row, n_cols - 1, _SAKOR_CONTRACT_HEADER)
    contract_header.font = Font(name=THAI_FONT, size=12, bold=True)
    contract_header.fill = _HEADER_FILL
    contract_header.alignment = Alignment(horizontal="center", wrap_text=True)
    contract_header.border = _BORDER
    ws.cell(row, n_cols).border = _BORDER
    row += 1

    po_groups = df.groupby(
        ["PONO", "ISSUEDATETIME", "VendorLabel", "PurchaseCategory"],
        sort=False,
    )
    po_agg = po_groups.agg(
        ItemCount=("STOCKCODE", "nunique"),
        NetTotal=("NetAmt", "sum"),
        FirstItemName=("ItemLabel", "first"),
    ).reset_index()
    po_agg = po_agg.sort_values("ISSUEDATETIME", kind="stable").reset_index(drop=True)

    for i, r in po_agg.iterrows():
        if r["ItemCount"] > 1:
            work_desc = f"{r['PurchaseCategory']} {r['ItemCount']} รายการ"
        else:
            work_desc = r["FirstItemName"]

        # "วิธีซื้อหรือจ้าง" กำหนดจากวงเงินที่ตกลงซื้อ/จ้างตามเกณฑ์ที่หน่วยงานระบุ ไม่ใช้
        # SKPO.PURCHASETYPECODE เพราะพบว่าไม่น่าเชื่อถือ (บางค่าเป็นชื่อหน่วยงานแทนวิธีจัดซื้อจริง)
        net_total = float(r["NetTotal"])
        purchase_method = "e-bidding" if net_total >= _SAKOR_EBIDDING_THRESHOLD else "วิธีเฉพาะเจาะจง"

        values = [
            i + 1, work_desc, net_total, net_total,
            purchase_method, r["VendorLabel"], net_total, r["VendorLabel"],
            net_total, _SAKOR_REASON_TEXT,
            r["PONO"], _to_thai_date_short(r["ISSUEDATETIME"]),
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row, col_idx, value)
            cell.font = Font(name=THAI_FONT, size=11)
            cell.border = _BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if col_idx in (3, 4, 7, 9):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in (1, 11, 12):
                cell.alignment = Alignment(horizontal="center", vertical="center")
        row += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
