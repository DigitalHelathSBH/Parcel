"""
เว็บแอปสำหรับหน่วยงาน - กรอกฟอร์มเลือกฝ่าย/แผนก/งาน และวันที่
เพื่อสร้างรายงานตรวจครุภัณฑ์ประจำปี (ที่กรองของตัดจำหน่ายแล้วออกให้อัตโนมัติ)

วิธีรัน:
    python app.py
แล้วเปิดเบราว์เซอร์ไปที่ http://localhost:5000

หมายเหตุ: ค่าเริ่มต้นรันแบบ localhost เท่านั้น (เข้าได้เฉพาะเครื่องนี้)
ถ้าต้องการให้เครื่องอื่นในหน่วยงานเข้าถึงผ่านเครือข่ายภายในได้ ต้องเปลี่ยน host
เป็น "0.0.0.0" ด้านล่าง - ควรปรึกษาฝ่าย IT ก่อน เพราะจะเปิดพอร์ตนี้ให้เครื่องอื่น
ในเครือข่ายเรียกเข้ามาได้ และควรพิจารณาเพิ่มการยืนยันตัวตนก่อนใช้งานจริง
"""
from datetime import date
from io import BytesIO

from flask import Flask, jsonify, render_template, request, send_file

import depreciation_builder as db
import report_builder as rb

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html", today=date.today().isoformat())


@app.route("/depreciation")
def depreciation_page():
    today = date.today()
    return render_template(
        "depreciation.html", today_year=today.year, today_month=today.month
    )


@app.route("/api/divisions")
def api_divisions():
    df = rb.list_divisions()
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/depts")
def api_depts():
    division = request.args.get("division", "")
    if not division:
        return jsonify([])
    df = rb.list_depts(division)
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/sections")
def api_sections():
    division = request.args.get("division", "")
    dept = request.args.get("dept", "")
    if not division or not dept:
        return jsonify([])
    df = rb.list_sections(division, dept)
    return jsonify(df.to_dict(orient="records"))


@app.route("/generate", methods=["POST"])
def generate():
    as_of_date = request.form.get("as_of_date") or date.today().isoformat()
    fiscal_year = request.form.get("fiscal_year") or ""
    fiscal_period = request.form.get("fiscal_period") or ""
    division = request.form.get("division") or None
    dept = request.form.get("dept") or None
    section = request.form.get("section") or None
    apply_depre_filter = request.form.get("apply_depre_filter") == "on"

    df = rb.fetch_asset_data(
        as_of_date,
        division=division,
        dept=dept,
        section=section,
        apply_depre_filter=apply_depre_filter,
    )

    if df.empty:
        return "ไม่พบรายการครุภัณฑ์ตามเงื่อนไขที่เลือก กรุณาลองเงื่อนไขอื่น", 400

    excel_bytes = rb.build_excel_bytes(df, fiscal_year, fiscal_period, as_of_date)

    filename = f"asset_annual_inspection_{as_of_date}.xlsx"
    return send_file(
        BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _read_depre_form():
    year = int(request.form.get("year"))
    month = int(request.form.get("month"))
    division = request.form.get("division") or None
    dept = request.form.get("dept") or None
    section = request.form.get("section") or None
    return year, month, division, dept, section


@app.route("/generate_depreciation_detail", methods=["POST"])
def generate_depreciation_detail():
    year, month, division, dept, section = _read_depre_form()
    df = db.fetch_depreciation_data(year, month, division=division, dept=dept, section=section)
    if df.empty:
        return "ไม่พบข้อมูลตามเงื่อนไขที่เลือก กรุณาลองเงื่อนไขอื่น", 400
    excel_bytes = db.build_detail_excel_bytes(df, year, month)
    filename = f"depre_detail_by_source_{year}-{month:02d}.xlsx"
    return send_file(
        BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/generate_depreciation_summary", methods=["POST"])
def generate_depreciation_summary():
    year, month, division, dept, section = _read_depre_form()
    df = db.fetch_depreciation_data(year, month, division=division, dept=dept, section=section)
    if df.empty:
        return "ไม่พบข้อมูลตามเงื่อนไขที่เลือก กรุณาลองเงื่อนไขอื่น", 400
    excel_bytes = db.build_summary_excel_bytes(df, year, month)
    filename = f"depre_summary_by_division_{year}-{month:02d}.xlsx"
    return send_file(
        BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
