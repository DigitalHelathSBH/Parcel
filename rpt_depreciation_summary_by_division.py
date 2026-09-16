# %%
"""
รายงานค่าเสื่อมราคา สรุปตามกลุ่มงาน
(ทดแทน AS_MAS_รายงานค่าเสื่อมราคา_สรุปตามกลุ่มงาน_ปร.rpt)

จัดกลุ่มตามฝ่าย/หน่วยงาน (LOCATEDIVISION) แล้วตามหมวดครุภัณฑ์ (ARTICLEGROUP)
แสดงเฉพาะยอดรวม ไม่มีรายบรรทัดทรัพย์สินแต่ละชิ้น
"""
import depreciation_builder as db

# ==== พารามิเตอร์รายงาน ====
YEAR = 2026     # ปี ค.ศ. ของเดือนที่ต้องการดู (เช่น สิงหาคม 2569 = 2026)
MONTH = 8       # เดือนปฏิทิน 1-12 (เช่น 8 = สิงหาคม)
DIVISION = None
DEPT = None
SECTION = None

OUTPUT_PATH = "depre_summary_by_division.xlsx"

# %%
df = db.fetch_depreciation_data(YEAR, MONTH, division=DIVISION, dept=DEPT, section=SECTION)
df

# %%
print(f"จำนวนรายการ: {len(df)}")
print(f"มูลค่ารวม: {df['TotalValue'].sum():,.2f} บาท")

# %%
excel_bytes = db.build_summary_excel_bytes(df, YEAR, MONTH)
with open(OUTPUT_PATH, "wb") as f:
    f.write(excel_bytes)
print(f"บันทึกไฟล์แล้ว: {OUTPUT_PATH}")
