# %%
"""
รายงานตรวจครุภัณฑ์ประจำปี (เรียงตามรหัส) - เวอร์ชัน CLI
สำหรับรันเองแบบไม่ผ่านหน้าเว็บ (ดูหน้าเว็บได้ที่ app.py)
"""
import report_builder as rb

# ==== พารามิเตอร์รายงาน (แก้ตรงนี้ก่อนรันแต่ละครั้ง) ====
AS_OF_DATE = "2026-09-30"   # แสดงเฉพาะครุภัณฑ์ที่ "ได้มา" ไม่เกินวันที่นี้
FISCAL_YEAR = "2569"        # ปี (พ.ศ.) ไว้แสดงในหัวรายงานเท่านั้น
FISCAL_PERIOD = "4"         # งวด ไว้แสดงในหัวรายงานเท่านั้น
DIVISION = None             # ระบุรหัสฝ่าย เช่น "102" หรือปล่อย None = ทุกฝ่าย
DEPT = None                 # ระบุรหัสแผนก หรือปล่อย None = ทุกแผนก
SECTION = None              # ระบุรหัสงาน หรือปล่อย None = ทุกงาน
APPLY_DEPRE_FILTER = False  # True = เฉพาะที่ขึ้นบัญชีสินทรัพย์ (DEPREGROUP=1, DEPREMETHOD=1)

OUTPUT_PATH = "asset_annual_inspection.xlsx"

# %%
df = rb.fetch_asset_data(
    AS_OF_DATE, division=DIVISION, dept=DEPT, section=SECTION,
    apply_depre_filter=APPLY_DEPRE_FILTER,
)
df

# %%
print(f"จำนวนรายการทั้งหมด: {len(df)}")
print(f"มูลค่ารวม: {df['PRICE'].sum():,.2f} บาท")

# %%
excel_bytes = rb.build_excel_bytes(df, FISCAL_YEAR, FISCAL_PERIOD, AS_OF_DATE)
with open(OUTPUT_PATH, "wb") as f:
    f.write(excel_bytes)
print(f"บันทึกไฟล์แล้ว: {OUTPUT_PATH}")
