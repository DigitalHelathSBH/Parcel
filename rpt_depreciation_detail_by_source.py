# %%
"""
รายงานค่าเสื่อมราคา รายละเอียด ตามแหล่งเงิน
(ทดแทน AS_MAS_รายงานค่าเสื่อมราคารายละเอียด_ตามแหล่ง.rpt)

จัดกลุ่มตามแหล่งเงิน (PURCHASEBUDGETCODE) แล้วตามหมวดครุภัณฑ์ (ARTICLEGROUP)
สูตรตรวจสอบกับไฟล์ตัวอย่าง .xls แล้ว ตรงกันทุกตัวเลข
"""
import depreciation_builder as db

# ==== พารามิเตอร์รายงาน ====
YEAR = 2026     # ปี ค.ศ. ของเดือนที่ต้องการดู (เช่น สิงหาคม 2569 = 2026)
MONTH = 8       # เดือนปฏิทิน 1-12 (เช่น 8 = สิงหาคม)
DIVISION = None
DEPT = None
SECTION = None

OUTPUT_PATH = "depre_detail_by_source.xlsx"

# %%
df = db.fetch_depreciation_data(YEAR, MONTH, division=DIVISION, dept=DEPT, section=SECTION)
df

# %%
print(f"จำนวนรายการ: {len(df)}")
print(f"มูลค่ารวม: {df['TotalValue'].sum():,.2f} บาท")
print(f"ค่าเสื่อมสะสม: {df['AccumDepre'].sum():,.2f} บาท")
print(f"มูลค่าสุทธิ: {df['NetValue'].sum():,.2f} บาท")

# %%
excel_bytes = db.build_detail_excel_bytes(df, YEAR, MONTH)
with open(OUTPUT_PATH, "wb") as f:
    f.write(excel_bytes)
print(f"บันทึกไฟล์แล้ว: {OUTPUT_PATH}")
