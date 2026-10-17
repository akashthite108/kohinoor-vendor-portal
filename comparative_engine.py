from datetime import datetime
import glob
import os
import openpyxl
import pandas as pd

def clean_procurement_dataframe(df):
    """Standardizes column names and cleans numeric NaNs to prevent crashes."""
    if df.empty:
        return df
        
    # 1. Standardize columns: strip whitespace, Title Case, replace spaces with underscores
    df.columns = [str(c).strip().title().replace(' ', '_') for c in df.columns]
    
    # 2. Force conversion of numeric columns to prevent Jinja format crashes
    numeric_columns = ['Rate_R0', 'Rate_R1', 'Quantity', 'Basic_Amount', 'Percentage_Difference_From_L1']
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
    return df

def generate_comparative_statement(rfq_code_target):
    template_path = "Comparision Format Sheet.xlsx"
    submissions_dir = "vendor_filled_slips"

    if not os.path.exists(template_path):
        print(f"Error: Template '{template_path}' not found.")
        return

    # Strictly match files by unique RFQ tracking code
    pattern = os.path.join(submissions_dir, f"*_{rfq_code_target}_*.xlsx")
    vendor_files = glob.glob(pattern)

    if not vendor_files:
        print(f"No vendor submissions found for RFQ Code: {rfq_code_target}")
        return

    if len(vendor_files) > 5:
        print(
            f"Warning: Found {len(vendor_files)} vendors. Maximum supported is 5."
            " Taking first 5."
        )
        vendor_files = vendor_files[:5]

    print(
        f"Found {len(vendor_files)} vendor submission(s) for RFQ:"
        f" {rfq_code_target}"
    )

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    project_name = ""
    tower_name = ""

    for idx, file_path in enumerate(vendor_files):
        # Vendor block columns: Vendor A (Cols E,F,G), Vendor B (Cols H,I,J), etc.
        start_col = 5 + (idx * 3)
        v_wb = None
        try:
            v_wb = openpyxl.load_workbook(file_path, data_only=True)
            v_ws = v_wb.active

            vendor_name = v_ws["F4"].value
            project_name = v_ws["B4"].value
            tower_name = v_ws["B5"].value

            print(
                f"Mapping Vendor {idx+1}: {vendor_name} -> Column block"
                f" {start_col}-{start_col+2}"
            )

            # Vendor Name Header (Row 4)
            ws.cell(row=4, column=start_col, value=vendor_name)

            # Read BOQ Items (Rows 11, 12, 13 in vendor slips mapped to 7, 8, 9 in comparison sheet)
            for i in range(3):
                v_row = 11 + i
                target_row = 7 + i

                desc = v_ws.cell(row=v_row, column=2).value
                qty = v_ws.cell(row=v_row, column=4).value
                unit = v_ws.cell(row=v_row, column=5).value
                brand = v_ws.cell(row=v_row, column=6).value
                rate_r0 = v_ws.cell(row=v_row, column=7).value or 0
                rate_r1 = v_ws.cell(row=v_row, column=8).value or rate_r0

                if desc:
                    ws.cell(row=target_row, column=2, value=desc)
                    ws.cell(row=target_row, column=3, value=qty)
                    ws.cell(row=target_row, column=4, value=unit)

                    # Brand Offered (Row 5)
                    ws.cell(row=5, column=start_col, value=brand)
                    # Initial Quoted Rate R0
                    ws.cell(row=target_row, column=start_col, value=rate_r0)
                    # Negotiated Rate R1
                    ws.cell(row=target_row, column=start_col + 1, value=rate_r1)

            # Commercial Terms Mapping
            comm_rows = {
                11: 'E18',  # Transport -> Row 11
                12: 'E19',  # GST -> Row 12
                14: 'E27',  # Earlier Supplied -> Row 14
                15: 'E21',  # Loading -> Row 15
                16: 'E23',  # Insurance -> Row 16
                17: 'E22',  # Unloading -> Row 17
                18: 'E24',  # Erection / Commissioning / Installation -> Row 18
                19: 'E25',  # Warranty -> Row 19
                20: 'E26',  # Guarantee -> Row 20
                21: 'E27',  # Earlier Supplied -> Row 21
                22: 'E28',  # Delivery Timeline -> Row 22
                23: 'E29',  # Payment Terms -> Row 23
                24: 'E30',  # Price Validity -> Row 24
            }

            for comp_row, v_cell in comm_rows.items():
                term_val = v_ws[v_cell].value
                ws.cell(row=comp_row, column=start_col, value=term_val)
        finally:
            if v_wb:
                v_wb.close()

    # Update Project Metadata Header
    ws["B2"] = project_name
    ws["B3"] = tower_name
    ws["L2"] = datetime.now().strftime("%Y-%m-%d")
    ws["L3"] = "Purchase Dept (Automated)"

    # Save unique comparative statement named after material/RFQ code
    output_filename = f"Kohinoor_Comparative_Statement_{rfq_code_target}.xlsx"
    wb.save(output_filename)
    wb.close()
    print(
        f"\nSuccess! Comparative Statement generated: '{output_filename}' with"
        f" {len(vendor_files)} vendor(s)."
    )

if __name__ == "__main__":
    target_rfq = "KWV-T1-ACME-001"
    generate_comparative_statement(target_rfq)