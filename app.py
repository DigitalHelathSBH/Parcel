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
import os
from datetime import date
from io import BytesIO

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.middleware.proxy_fix import ProxyFix

import depreciation_builder as db
import disposal_builder as disp
import low_value_builder as lv
import purchase_builder as pb
import report_builder as rb

app = Flask(__name__)
# behind nginx we're reverse-proxied under a subpath (e.g. /parcel/); honor
# X-Forwarded-Prefix (and the usual X-Forwarded-*) so url_for() builds correct links
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)


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


@app.route("/low_value")
def low_value_page():
    return render_template("low_value.html", today=date.today().isoformat())


@app.route("/generate_low_value", methods=["POST"])
def generate_low_value():
    as_of_date = request.form.get("as_of_date") or date.today().isoformat()
    division = request.form.get("division") or None
    dept = request.form.get("dept") or None
    section = request.form.get("section") or None

    df = lv.fetch_low_value_data(as_of_date, division=division, dept=dept, section=section)
    if df.empty:
        return "ไม่พบครุภัณฑ์ต่ำกว่าเกณฑ์ตามเงื่อนไขที่เลือก กรุณาลองเงื่อนไขอื่น", 400

    excel_bytes = lv.build_excel_bytes(df, as_of_date)
    filename = f"low_value_assets_{as_of_date}.xlsx"
    return send_file(
        BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/purchases")
def purchases_page():
    today = date.today()
    return render_template("purchases.html", today_year=today.year, today_month=today.month)


@app.route("/generate_purchases", methods=["POST"])
def generate_purchases():
    year = int(request.form.get("year"))
    month = int(request.form.get("month"))

    df = pb.fetch_purchase_data(year, month)
    if df.empty:
        return "ไม่พบรายการจัดซื้อในเดือนที่เลือก กรุณาลองเดือนอื่น", 400

    excel_bytes = pb.build_excel_bytes(df, year, month)
    filename = f"purchases_{year}-{month:02d}.xlsx"
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


@app.route("/generate_depreciation_source_summary", methods=["POST"])
def generate_depreciation_source_summary():
    year, month, division, dept, section = _read_depre_form()
    df = db.fetch_depreciation_data(year, month, division=division, dept=dept, section=section)
    if df.empty:
        return "ไม่พบข้อมูลตามเงื่อนไขที่เลือก กรุณาลองเงื่อนไขอื่น", 400
    excel_bytes = db.build_source_summary_excel_bytes(df, year, month)
    filename = f"depre_source_summary_{year}-{month:02d}.xlsx"
    return send_file(
        BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/disposal")
def disposal_page():
    today = date.today()
    return render_template("disposal.html", today_year=today.year, today_month=today.month)


@app.route("/generate_disposal", methods=["POST"])
def generate_disposal():
    year = int(request.form.get("year"))
    month = int(request.form.get("month"))
    division = request.form.get("division") or None
    dept = request.form.get("dept") or None
    section = request.form.get("section") or None

    df = disp.fetch_disposal_data(year, month, division=division, dept=dept, section=section)
    if df.empty:
        return "ไม่พบรายการตัดจำหน่ายตามเงื่อนไขที่เลือก กรุณาลองเงื่อนไขอื่น", 400

    excel_bytes = disp.build_excel_bytes(df, year, month)
    filename = f"disposal_detail_{year}-{month:02d}.xlsx"
    return send_file(
        BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="127.0.0.1", port=port, debug=False)
