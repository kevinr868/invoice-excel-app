import io
import json
import openpyxl
import streamlit as st
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from typing import List, Optional
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="PDF Invoice to Excel Converter", page_icon="📊", layout="wide")

st.title("📊 PDF Invoice to Formatted Excel Converter")
st.write("Upload a PDF invoice to automatically extract line items and download a styled Excel file.")

# --- API KEY CHECK ---
gemini_api_key = st.secrets.get("GEMINI_API_KEY", None)

if not gemini_api_key:
    with st.sidebar:
        st.header("🔑 Configuration")
        gemini_api_key = st.text_input("Enter Gemini API Key", type="password")
        st.markdown("[Get a Gemini API Key](https://aistudio.google.com/app/apikey)")

# --- PYDANTIC SCHEMAS ---
class InvoiceItem(BaseModel):
    item_no: Optional[int] = Field(description="Item or line number")
    material_no: Optional[str] = Field(description="Material/Part/SKU number")
    description: str = Field(description="Description of product or service")
    qty: float = Field(description="Quantity ordered or shipped")
    unit_price: float = Field(description="Unit price per item")
    origin: Optional[str] = Field(description="Country of origin")
    batch: Optional[str] = Field(description="Batch number or Serial Number")
    expiry: Optional[str] = Field(description="Expiration date if present")

class InvoiceData(BaseModel):
    invoice_number: str = Field(description="Invoice number")
    invoice_date: str = Field(description="Invoice date")
    due_date: Optional[str] = Field(description="Payment due date")
    payment_terms: Optional[str] = Field(description="Payment terms, e.g., 150 days")
    currency: Optional[str] = Field(description="Invoice currency, e.g., USD")
    customer: Optional[str] = Field(description="Customer name / Bill To")
    address: Optional[str] = Field(description="Billing address")
    country: Optional[str] = Field(description="Customer country")
    ship_to: Optional[str] = Field(description="Ship To location")
    po_number: Optional[str] = Field(description="Purchase Order number")
    order_number: Optional[str] = Field(description="Sales Order number and date")
    contract_mark: Optional[str] = Field(description="Contract number or Mark reference")
    freight: Optional[float] = Field(default=0.0, description="Freight cost / Flete")
    insurance: Optional[float] = Field(default=0.0, description="Insurance cost / Seguro")
    items: List[InvoiceItem] = Field(description="List of line items")

# --- EXCEL GENERATOR FUNCTION ---
def create_excel_workbook(data: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice Details"
    ws.views.sheetView[0].showGridLines = True

    font_fam = "Segoe UI"
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name=font_fam, size=11, bold=True, color="FFFFFF")
    section_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    section_font = Font(name=font_fam, size=11, bold=True, color="1F497D")
    title_font = Font(name=font_fam, size=16, bold=True, color="1F497D")
    label_font = Font(name=font_fam, size=10, bold=True, color="333333")
    data_font = Font(name=font_fam, size=10, color="000000")
    bold_data_font = Font(name=font_fam, size=10, bold=True, color="000000")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    top_thin_bottom_double = Border(top=Side(border_style="thin", color="000000"), bottom=Side(border_style="double", color="000000"))
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    # Title Block
    ws.merge_cells("A1:I1")
    ws["A1"] = "INVOICE DETAILS"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 25

    # Header Section
    ws.merge_cells("A3:I3")
    ws["A3"] = "HEADER INFORMATION"
    ws["A3"].font = section_font
    ws["A3"].fill = section_fill
    ws["A3"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[3].height = 20

    header_info = [
        ("Invoice Number:", data.get("invoice_number", ""), "Customer / Bill To:", data.get("customer", "")),
        ("Invoice Date:", data.get("invoice_date", ""), "Address:", data.get("address", "")),
        ("Due Date:", data.get("due_date", ""), "Country:", data.get("country", "")),
        ("Payment Terms:", data.get("payment_terms", ""), "Ship To:", data.get("ship_to", "")),
        ("Currency:", data.get("currency", "USD"), "PO Number:", data.get("po_number", "")),
        ("Order Number:", data.get("order_number", ""), "Contract / Mark:", data.get("contract_mark", ""))
    ]

    r = 4
    for label1, val1, label2, val2 in header_info:
        ws.cell(row=r, column=1, value=label1).font = label_font
        ws.cell(row=r, column=2, value=str(val1 or "")).font = data_font
        ws.cell(row=r, column=4, value=label2).font = label_font
        ws.cell(row=r, column=5, value=str(val2 or "")).font = data_font
        r += 1

    headers = [
        "Item", "Material No.", "Description", "Quantity", 
        f"Unit Price ({data.get('currency', 'USD')})", 
        f"Net Price ({data.get('currency', 'USD')})", 
        "Country of Origin", "Batch / SN", "Expiry Date"
    ]
    
    start_row_items = 11
    ws.row_dimensions[start_row_items].height = 24

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row_items, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if col_idx != 3 else "left", vertical="center", wrap_text=True)

    curr_r = start_row_items + 1
    items = data.get("items", [])
    
    for idx, item in enumerate(items):
        ws.cell(row=curr_r, column=1, value=item.get("item_no")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=2, value=str(item.get("material_no") or "")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=3, value=item.get("description", "")).alignment = Alignment(horizontal="left")
        
        c_qty = ws.cell(row=curr_r, column=4, value=item.get("qty", 0))
        c_qty.number_format = '#,##0'
        c_qty.alignment = Alignment(horizontal="right")
        
        c_uprice = ws.cell(row=curr_r, column=5, value=item.get("unit_price", 0.0))
        c_uprice.number_format = '$#,##0.00'
        c_uprice.alignment = Alignment(horizontal="right")
        
        c_net = ws.cell(row=curr_r, column=6, value=f"=D{curr_r}*E{curr_r}")
        c_net.number_format = '$#,##0.00'
        c_net.alignment = Alignment(horizontal="right")
        
        ws.cell(row=curr_r, column=7, value=str(item.get("origin") or "")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=8, value=str(item.get("batch") or "")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=9, value=str(item.get("expiry") or "")).alignment = Alignment(horizontal="center")
        
        for c in range(1, 10):
            cell = ws.cell(row=curr_r, column=c)
            cell.font = data_font
            cell.border = thin_border
            if idx % 2 == 1:
                cell.fill = zebra_fill
                
        ws.row_dimensions[curr_r].height = 20
        curr_r += 1

    curr_r += 1
    ws.cell(row=curr_r, column=5, value="FOB Subtotal:").font = bold_data_font
    ws.cell(row=curr_r, column=5).alignment = Alignment(horizontal="right")
    c_fob = ws.cell(row=curr_r, column=6, value=f"=SUM(F{start_row_items+1}:F{curr_r-2})")
    c_fob.font = bold_data_font
    c_fob.number_format = '$#,##0.00'
    c_fob.alignment = Alignment(horizontal="right")
    c_fob.border = Border(top=Side(border_style="thin", color="000000"))
    curr_r += 1

    ws.cell(row=curr_r, column=5, value="Freight (Flete):").font = data_font
    ws.cell(row=curr_r, column=5).alignment = Alignment(horizontal="right")
    c_freight = ws.cell(row=curr_r, column=6, value=data.get("freight", 0.0))
    c_freight.font = data_font
    c_freight.number_format = '$#,##0.00'
    c_freight.alignment = Alignment(horizontal="right")
    curr_r += 1

    ws.cell(row=curr_r, column=5, value="Insurance (Seguro):").font = data_font
    ws.cell(row=curr_r, column=5).alignment = Alignment(horizontal="right")
    c_ins = ws.cell(row=curr_r, column=6, value=data.get("insurance", 0.0))
    c_ins.font = data_font
    c_ins.number_format = '$#,##0.00'
    c_ins.alignment = Alignment(horizontal="right")
    curr_r += 1

    ws.cell(row=curr_r, column=5, value="Total CIP:").font = bold_data_font
    ws.cell(row=curr_r, column=5).alignment = Alignment(horizontal="right")
    c_cip = ws.cell(row=curr_r, column=6, value=f"=F{curr_r-3}+F{curr_r-2}+F{curr_r-1}")
    c_cip.font = Font(name=font_fam, size=11, bold=True, color="1F497D")
    c_cip.number_format = '$#,##0.00'
    c_cip.alignment = Alignment(horizontal="right")
    c_cip.border = top_thin_bottom_double
    ws.cell(row=curr_r, column=5).border = top_thin_bottom_double

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 3] or cell.column in [3, 5]:
                continue
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        
        if col_letter == 'C':
            ws.column_dimensions[col_letter].width = 42
        elif col_letter in ['E', 'F']:
            ws.column_dimensions[col_letter].width = 18
        else:
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- WORKFLOW ---
uploaded_file = st.file_uploader("Upload Invoice PDF", type=["pdf"])

if uploaded_file is not None:
    # Ensure key string is stripped of whitespace or quotes
    clean_key = gemini_api_key.strip().strip('"').strip("'") if gemini_api_key else ""
    
    if not clean_key:
        st.error("Please provide a valid Gemini API Key in the sidebar or Secrets.")
    else:
        if st.button("🚀 Process Invoice & Generate Excel", type="primary"):
            with st.spinner("Analyzing PDF with Gemini 1.5 Pro..."):
                try:
                    # Initialize Google GenAI Client
                    client = genai.Client(api_key=clean_key)
                    
                    pdf_bytes = uploaded_file.read()

                    # Call Gemini using standard Part.from_bytes
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[
                            types.Part.from_bytes(
                                data=pdf_bytes,
                                mime_type='application/pdf',
                            ),
                            "Analyze this invoice document carefully. Extract all header fields, "
                            "line items (including material numbers, batch numbers, origin, and expiry dates), "
                            "and financial totals (Freight, Insurance). Return exact structured JSON."
                        ],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=InvoiceData,
                            temperature=0.0
                        )
                    )

                    extracted_data = json.loads(response.text)
                    excel_data = create_excel_workbook(extracted_data)

                    st.success("Extraction Complete!")
                    st.download_button(
                        label="📥 Download Formatted Excel Workbook",
                        data=excel_data,
                        file_name=f"Invoice_{extracted_data.get('invoice_number', 'Data')}_Details.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")
