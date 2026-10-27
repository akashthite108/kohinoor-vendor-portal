@app.route('/generate-comparison', methods=['POST'])
def generate_comparison():
    template_path = TEMPLATE_SHEET
    if not os.path.exists(template_path):
        return "<script>alert('Error: Comparison format template sheet missing.'); window.location.href='/purchase-dashboard';</script>", 500

    pool_files = glob.glob(os.path.join(COMPARISON_POOL_DIR, '*.xlsx'))
    pool_files = [f for f in pool_files if not os.path.basename(f).startswith('Kohinoor_Comparative') and not os.path.basename(f).startswith('~$')]
    
    if len(pool_files) < 2:
        return "<script>alert('Strict Rule Error: Comparison requires a minimum of 2 vendors (maximum 5). Please add more vendors to the comparison pool.'); window.location.href='/purchase-dashboard';</script>", 400
    if len(pool_files) > 5:
        pool_files = pool_files[:5]

    comparison_title = request.form.get('comparison_title', 'General Civil Materials Procurement Comparison')
    timestamp_code = datetime.now().strftime('%Y%m%d-%H%M%S')
    unique_comp_code = f'COMP-{timestamp_code}'

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    def safe_set(row_num, col_num, val):
        cell = ws.cell(row=row_num, column=col_num)
        coord = cell.coordinate
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if coord in merged_range:
                    ws.cell(row=merged_range.min_row, column=merged_range.min_col, value=val)
                    return
        cell.value = val

    project_name = ''
    tower_name = ''
    target_rfq_code = 'COMPARATIVE'

    for idx, file_path in enumerate(pool_files):
        # Vendor block columns: Vendor A starts at Col F (6), Vendor B at Col I (9), etc.
        start_col = 6 + (idx * 3)
        v_wb = None
        try:
            v_wb = openpyxl.load_workbook(file_path, data_only=True)
            v_ws = v_wb.active

            vendor_name = v_ws['F4'].value
            project_name = v_ws['B4'].value
            tower_name = v_ws['B5'].value
            target_rfq_code = v_ws['F5'].value or 'COMPARATIVE'

            # Set Vendor Name in Row 4 safely
            safe_set(4, start_col, vendor_name)

            boq_rows = get_boq_rows(v_ws)
            for i, v_row in enumerate(boq_rows):
                target_row = 6 + i  # BOQ items in new template are rows 6 to 10
                if target_row > 10:
                    break

                desc = v_ws.cell(row=v_row, column=2).value
                qty = v_ws.cell(row=v_row, column=4).value
                unit = v_ws.cell(row=v_row, column=5).value
                brand = v_ws.cell(row=v_row, column=6).value
                rate_r0 = v_ws.cell(row=v_row, column=7).value or 0
                rate_r1 = v_ws.cell(row=v_row, column=8).value or rate_r0

                if desc and str(desc).strip():
                    safe_set(target_row, 2, desc)
                    safe_set(target_row, 3, brand)   # Make / Brand Offered in Column C
                    safe_set(target_row, 4, qty)     # Qty in Column D
                    safe_set(target_row, 5, unit)    # Unit in Column E
                    safe_set(target_row, start_col, rate_r0)      # Rate R0
                    safe_set(target_row, start_col + 1, rate_r1)  # Rate R1

            # Commercial terms mapping for new template rows
            comm_rows = {
                12: 'E22',  # Transport
                13: 'E23',  # GST
                15: 'E31',  # Earlier Supplied
                16: 'E25',  # Loading
                17: 'E27',  # Insurance
                18: 'E26',  # Unloading
                19: 'E28',  # Erection
                20: 'E29',  # Warranty
                21: 'E30',  # Guarantee
                23: 'E32',  # Delivery
                24: 'E33',  # Payment Terms
                25: 'E34',  # Price Validity
                26: 'E35',  # Remark
            }

            for comp_row, v_cell in comm_rows.items():
                term_val = v_ws[v_cell].value
                safe_set(comp_row, start_col, term_val)
        finally:
            if v_wb:
                v_wb.close()

    ws['B2'] = project_name
    ws['B3'] = tower_name
    ws['M2'] = datetime.now().strftime('%Y-%m-%d')
    ws['M3'] = 'Purchase Dept (Automated)'

    safe_target_rfq = sanitize_filename(target_rfq_code)
    version_num = 1
    while True:
        output_filename = f'Kohinoor_Comparative_{safe_target_rfq}_{unique_comp_code}_V{version_num}.xlsx'
        output_path = os.path.join(COMPARISON_POOL_DIR, output_filename)
        if not os.path.exists(output_path):
            break
        version_num += 1

    wb.save(output_path)
    wb.close()
    rebuild_master_database()
    return send_file(output_path, as_attachment=True)