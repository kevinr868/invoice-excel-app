import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import json
import openai  # or google-generativeai

# --- STREAMLIT UI CONFIG ---
st.set_page_config(page_title="Invoice to Excel Converter", layout="centered")
st.title("📄 PDF Invoice to Formatted Excel Converter")
st.write("Upload a PDF invoice to extract line items and download a styled Excel workbook.")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

# --- EXCEL GENERATOR FUNCTION ---
def create_excel_workbook(extracted_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice Details"
    ws.views.sheetView[0].showGridLines = True

    # Styling Setup
    font_fam = "Segoe UI"
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name=font_fam, size=11, bold=True, color="FFFFFF")
    section_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    section_font = Font(name=font_fam, size=11, bold=True, color="1F497D")
    label_font = Font(name=font_fam, size=10, bold=True, color="333333")
    data_font = Font(name=font_fam, size=10, color="000000")
    bold_data_font = Font(name=font_fam, size=10, bold=True, color="000000")
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    # Header Section
    ws.merge_cells("A1:I1")
    ws["A1"] = "INVOICE DETAILS"
    ws["A1"].font = Font(name=font_fam, size=16, bold=True, color="1F497D")

    ws.merge_cells("A3:I3")
    ws["A3"] = "HEADER INFORMATION"
    ws["A3"].font = section_font
    ws["A3"].fill = section_fill

    header_info = [
        ("Invoice Number:", extracted_data.get("invoice_number", ""), "Customer / Bill To:", extracted_data.get("customer", "")),
        ("Invoice Date:", extracted_data.get("invoice_date", ""), "Address:", extracted_data.get("address", "")),
        ("Due Date:", extracted_data.get("due_date", ""), "Country:", extracted_data.get("country", "")),
        ("Payment Terms:", extracted_data.get("payment_terms", ""), "Ship To:", extracted_data.get("ship_to", "")),
        ("Currency:", extracted_data.get("currency", "USD"), "PO Number:", extracted_data.get("po_number", "")),
        ("Order Number:", extracted_data.get("order_number", ""), "Contract / Mark:", extracted_data.get("contract_mark", ""))
    ]

    r = 4
    for l1, v1, l2, v2 in header_info:
        ws.cell(row=r, column=1, value=l1).font = label_font
        ws.cell(row=r, column=2, value=v1).font = data_font
        ws.cell(row=r, column=4, value=l2).font = label_font
        ws.cell(row=r, column=5, value=v2).font = data_font
        r += 1

    # Table Setup
    headers = ["Item", "Material No.", "Description", "Quantity", "Unit Price (USD)", "Net Price (USD)", "Country of Origin", "Batch / SN", "Expiry Date"]
    start_row = 11
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if col_idx != 3 else "left", vertical="center")

    # Data Rows
    curr_r = start_row + 1
    for idx, item in enumerate(extracted_data.get("items", [])):
        ws.cell(row=curr_r, column=1, value=item.get("item_no")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=2, value=item.get("material_no")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=3, value=item.get("description"))
        
        c_qty = ws.cell(row=curr_r, column=4, value=item.get("qty"))
        c_qty.number_format = '#,##0'
        
        c_price = ws.cell(row=curr_r, column=5, value=item.get("unit_price"))
        c_price.number_format = '$#,##0.00'
        
        # Formulas
        c_net = ws.cell(row=curr_r, column=6, value=f"=D{curr_r}*E{curr_r}")
        c_net.number_format = '$#,##0.00'

        ws.cell(row=curr_r, column=7, value=item.get("origin")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=8, value=item.get("batch")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=9, value=item.get("expiry")).alignment = Alignment(horizontal="center")

        for c in range(1, 10):
            cell = ws.cell(row=curr_r, column=c)
            cell.font = data_font
            cell.border = thin_border
            if idx % 2 == 1:
                cell.fill = zebra_fill
        curr_r += 1

    # Financial Totals
    curr_r += 1
    ws.cell(row=curr_r, column=5, value="FOB Subtotal:").font = bold_data_font
    c_fob = ws.cell(row=curr_r, column=6, value=f"=SUM(F{start_row+1}:F{curr_r-2})")
    c_fob.font = bold_data_font
    c_fob.number_format = '$#,##0.00'

    curr_r += 1
    ws.cell(row=curr_r, column=5, value="Freight (Flete):").font = data_font
    c_fr = ws.cell(row=curr_r, column=6, value=extracted_data.get("freight", 0))
    c_fr.number_format = '$#,##0.00'

    curr_r += 1
    ws.cell(row=curr_r, column=5, value="Insurance (Seguro):").font = data_font
    c_ins = ws.cell(row=curr_r, column=6, value=extracted_data.get("insurance", 0))
    c_ins.number_format = '$#,##0.00'

    curr_r += 1
    ws.cell(row=curr_r, column=5, value="Total CIP:").font = Font(name=font_fam, size=11, bold=True, color="1F497D")
    c_cip = ws.cell(row=curr_r, column=6, value=f"=F{curr_r-3}+F{curr_r-2}+F{curr_r-1}")
    c_cip.font = Font(name=font_fam, size=11, bold=True, color="1F497D")
    c_cip.number_format = '$#,##0.00'

    # Auto-fit columns
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'C':
            ws.column_dimensions[col_letter].width = 42
        elif col_letter in ['E', 'F']:
            ws.column_dimensions[col_letter].width = 18
        else:
            ws.column_dimensions[col_letter].width = 16

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- RUN CONVERSION ---
if uploaded_file is not None:
    if st.button("Extract Data & Generate Excel"):
        with st.spinner("Processing PDF and extracting fields via AI..."):
            
            # Simulated AI extraction call output schema
            # In production, pass uploaded_file bytes to OpenAI / Gemini API here
            extracted_data = {
                "invoice_number": "4553134879",
                "invoice_date": "18.09.2026",
                "due_date": "15.02.2027",
                "payment_terms": "150 days",
                "customer": "BRYDEN PI LIMITED",
                "freight": 379.39,
                "insurance": 104.33,
                "items": [
                    {"item_no": 2, "material_no": "03246353001", "description": "Cartridge CL", "qty": 6, "unit_price": 360.00, "origin": "Japan", "batch": "IFM", "expiry": "28.07.2027"}
                ]
            }

            excel_bytes = create_excel_workbook(extracted_data)

            st.success("Excel generated successfully!")
            st.download_button(
                label="📥 Download Excel File",
                data=excel_bytes,
                file_name=f"Invoice_{extracted_data.get('invoice_number', 'Output')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
