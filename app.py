from datetime import datetime
import os
import openpyxl
import pandas as pd
from flask import Flask, redirect, render_template, request, url_for
import re

app = Flask(__name__)
LOG_FILE = "master_staging_log.csv"
SUBMISSIONS_DIR = "vendor_filled_slips"

# Official Indian GSTIN Regex Pattern
GSTIN_REGEX = re.compile(
    r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
)


def init_db():
  os.makedirs(SUBMISSIONS_DIR, exist_ok=True)
  if not os.path.exists(LOG_FILE):
    df = pd.DataFrame(columns=[
        "Timestamp",
        "RFQ_Code",
        "Project_Name",
        "Tower",
        "Vendor_Legal_Name",
        "GSTIN",
        "Total_Basic",
        "Transport",
        "GST",
        "Final_Value",
        "Delivery",
        "Payment",
    ])
    try:
      df.to_csv(LOG_FILE, index=False)
    except Exception:
      pass


@app.route("/")
def index():
  return render_template("index.html")


@app.route("/submit", methods=["POST"])
def submit():
  project = request.form.get("project_name")
  tower = request.form.get("tower")
  enquiry_date = request.form.get("enquiry_date")
  buyer_org = request.form.get("buyer_org")
  buyer_contact = request.form.get("buyer_contact")

  vendor = request.form.get("vendor_name")
  rfq_code = request.form.get("rfq_code")
  quotation_date = request.form.get("quotation_date")
  gstin = request.form.get("gstin").strip().upper()
  vendor_contact = request.form.get("vendor_contact")

  # STRICT GSTIN VALIDATION GUARD
  if not GSTIN_REGEX.match(gstin):
    return (
        "<h3>Security Error: Invalid GSTIN format entered. A valid 15-character"
        " GSTIN is required (e.g., 27AAAAA0000A1Z5).</h3><a"
        " href='/'>Back</a>"
    )

  descriptions = request.form.getlist("desc[]")
  quantities = request.form.getlist("qty[]")
  units = request.form.getlist("unit[]")
  brands = request.form.getlist("brand[]")
  rates = request.form.getlist("rate[]")

  comm_transport_term = request.form.get("comm_transport")
  gst_rate_sel = request.form.get("gst_rate")
  comm_gst_term = f"{gst_rate_sel}% GST"
  comm_loading = request.form.get("comm_loading")
  comm_unloading = request.form.get("comm_unloading")
  comm_insurance = request.form.get("comm_insurance")
  comm_erection = request.form.get("comm_erection")
  comm_warranty = request.form.get("comm_warranty")
  comm_guarantee = request.form.get("comm_guarantee")
  comm_supplied = request.form.get("comm_supplied")
  comm_delivery = request.form.get("comm_delivery")
  comm_payment = request.form.get("comm_payment")
  comm_validity = request.form.get("comm_validity")

  total_basic = float(request.form.get("hidden_basic") or 0)
  transport_cost = float(request.form.get("hidden_transport") or 0)
  gst_amount = float(request.form.get("hidden_gst") or 0)
  final_amount = float(request.form.get("hidden_final") or 0)

  if total_basic <= 0:
    return (
        "<h3>Error: Basic value cannot be zero. Please go back and fill valid"
        " quantities and rates.</h3><a href='/'>Back</a>"
    )

  timestamp_file = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
  timestamp_display = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

  template_path = "RFQ.xlsx"
  if os.path.exists(template_path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    ws["B4"] = project
    ws["B5"] = tower
    ws["B6"] = enquiry_date
    ws["B7"] = buyer_org
    ws["B8"] = buyer_contact

    ws["F4"] = vendor
    ws["F5"] = rfq_code
    ws["F6"] = quotation_date
    ws["F7"] = gstin
    ws["F8"] = vendor_contact

    start_row = 10
    for i in range(len(descriptions)):
      current_row = start_row + i
      desc = descriptions[i]
      qty = float(quantities[i] or 0)
      unit = units[i]
      brand = brands[i]
      rate = float(rates[i] or 0)
      amount = qty * rate

      ws.cell(row=current_row, column=1, value=i + 1)
      ws.cell(row=current_row, column=2, value=desc)
      ws.cell(row=current_row, column=4, value=qty)
      ws.cell(row=current_row, column=5, value=unit)
      ws.cell(row=current_row, column=6, value=brand)
      ws.cell(row=current_row, column=7, value=rate)
      ws.cell(row=current_row, column=8, value=amount)

    total_row = start_row + len(descriptions)
    ws.cell(row=total_row, column=8, value=total_basic)

    ws["E17"] = "Included in BOQ"
    ws["H17"] = total_basic

    ws["E18"] = comm_transport_term
    ws["H18"] = transport_cost

    ws["E19"] = comm_gst_term
    ws["H19"] = gst_amount

    ws["E20"] = f"Calculated Total ({gst_rate_sel}% GST)"
    ws["H20"] = final_amount

    commercial_mapping = {
        21: comm_loading,
        22: comm_unloading,
        23: comm_insurance,
        24: comm_erection,
        25: comm_warranty,
        26: comm_guarantee,
        27: comm_supplied,
        28: comm_delivery,
        29: comm_payment,
        30: comm_validity,
    }
    for row_idx, term_value in commercial_mapping.items():
      ws.cell(row=row_idx, column=5, value=term_value)

    # Sanitize vendor company name for filename
    safe_vendor_name = "".join(
        c if c.isalnum() else "_" for c in vendor
    ).strip("_")
    output_filename = f"RFQ_{safe_vendor_name}_{rfq_code}_{timestamp_file}.xlsx"
    output_slip_path = os.path.join(SUBMISSIONS_DIR, output_filename)
    wb.save(output_slip_path)

  new_row = {
      "Timestamp": timestamp_display,
      "RFQ_Code": rfq_code,
      "Project_Name": project,
      "Tower": tower,
      "Vendor_Legal_Name": vendor,
      "GSTIN": gstin,
      "Total_Basic": total_basic,
      "Transport": transport_cost,
      "GST": gst_amount,
      "Final_Value": final_amount,
      "Delivery": comm_delivery,
      "Payment": comm_payment,
  }
  try:
    if os.path.exists(LOG_FILE):
      df = pd.read_csv(LOG_FILE)
      df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
      df.to_csv(LOG_FILE, index=False)
    else:
      pd.DataFrame([new_row]).to_csv(LOG_FILE, index=False)
  except PermissionError:
    pass

  return (
      "<h2 style='text-align:center; color:green; margin-top:50px;'>Official"
      " RFQ Quotation Submitted Successfully! Verified & logged under vendor"
      " name. <a href='/'>Submit Another</a></h2>"
  )


if __name__ == "__main__":
  init_db()
  port = int(os.environ.get("PORT", 8080))
  app.run(host='0.0.0.0', port=port)