# app.py
import streamlit as st
import pandas as pd
import zipfile
import pdfplumber
import re
import os
import io
from openpyxl import load_workbook
from zipfile import ZipFile
import base64

st.title("Invoice Processing Tool with ZIP or Single PDF Upload")

zip_file = st.file_uploader("Upload ZIP of Invoices (Optional)", type=["zip"])
pdf_file = st.file_uploader("Or Upload a Single Invoice PDF", type=["pdf"])
paf_file = st.file_uploader("Upload PAF Excel", type=["xlsx"])
template_file = st.file_uploader("Upload Invoice Template Excel", type=["xlsx"])

temp_folder = "temp_excels"
final_folder = "final_outputs"
final_invoice_folder = "final_invoice_output"
summary_file = "Summary_Report.xlsx"

for folder in [temp_folder, final_folder, final_invoice_folder]:
    os.makedirs(folder, exist_ok=True)

line_pattern = re.compile(
    r"^\s*(\d+)\s+"  # pattern for extracting invoice data
    r"(\S+)\s+" 
    r"(.+?)\s+" 
    r"(\d+(?:\.?\d*)?(?:EA|PAC))\s+" 
    r"([\d\.]+)\s+" 
    r"([\d\.]+)\s*$"
)

def extract_invoice_data(file):
    # Extract invoice data from the PDF
    products = []
    invoice_number = pic_number = freight_charges = gst_amount = total_tax_included = order_number = ship_to_address = None
    skip = False

    with pdfplumber.open(file) as pdf:
        first_text = pdf.pages[0].extract_text()

        inv_m = re.search(r"(INV\d{6})", first_text)
        pic_m = re.search(r"(PIC\d{6})", first_text)
        invoice_number = inv_m.group(1) if inv_m else None
        pic_number = pic_m.group(1) if pic_m else None

        order_number = ""
        for line in first_text.splitlines():
            match = re.search(r"\b(RGRHO\w+|CCAO\w+)\b", line)
            if match:
                order_number = match.group(1)
                break

        if order_number.startswith("CCAO"):
            skip = True

        ship_to_lines = re.findall(r"Ship To\s*\n(.*?)\n(.*?)\n", first_text, re.DOTALL)
        ship_to_address = "\n".join(ship_to_lines[0]) if ship_to_lines else ""

        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            try:
                header_idx = next(i for i, line in enumerate(lines) if line.strip().startswith("SNo.") and "Extension Cost" in line)
            except StopIteration:
                continue

            for line in lines[header_idx+1:]:
                if line.strip().startswith("Page"):
                    break
                m = line_pattern.match(line)
                if m:
                    quantity_full = m.group(4)
                    quantity_number = re.match(r"([\d\.]+)", quantity_full).group(1)
                    quantity_unit = re.search(r"(EA|PAC)", quantity_full).group(1)
                    products.append({
                        "SNo": int(m.group(1)),
                        "Product": m.group(2),
                        "Description": m.group(3).strip(),
                        "Quantity": float(quantity_number),
                        "Unit": quantity_unit,
                        "Gross Price": float(m.group(5)),
                        "Extension Cost": float(m.group(6))
                    })

        last_text = pdf.pages[-1].extract_text()
        f_m = re.search(r"Freight Charges\s+([\d,]+\.\d{2})", last_text)
        g_m = re.search(r"GST/HST Amount\s+([\d,]+\.\d{2})", last_text)
        t_m = re.search(r"TOTAL TAX INCLUDED\s+([\d,]+\.\d{2})", last_text, re.IGNORECASE)

        freight_charges = float(f_m.group(1).replace(",", "")) if f_m else 0
        gst_amount = float(g_m.group(1).replace(",", "")) if g_m else 0
        total_tax_included = float(t_m.group(1).replace(",", "")) if t_m else 0

    df = pd.DataFrame(products)
    return df, invoice_number, pic_number, freight_charges, gst_amount, total_tax_included, order_number, ship_to_address, skip

if (zip_file or pdf_file) and paf_file and template_file:
    if st.button("Process Invoices"):

        for folder in [temp_folder, final_folder, final_invoice_folder]:
            for f in os.listdir(folder):
                os.remove(os.path.join(folder, f))

        with open("uploaded_paf.xlsx", "wb") as f:
            f.write(paf_file.getbuffer())
        with open("uploaded_template.xlsx", "wb") as f:
            f.write(template_file.getbuffer())

        paf_df = pd.read_excel("uploaded_paf.xlsx")
        paf_df.columns = paf_df.columns.str.strip()
        paf_df = paf_df.drop_duplicates(subset=["Valiant/RGR SKU"])

        pdf_sources = []
        if zip_file:
            with open("uploaded_invoices.zip", "wb") as f:
                f.write(zip_file.getbuffer())
            with zipfile.ZipFile("uploaded_invoices.zip", 'r') as zip_ref:
                pdf_sources = [(f, zip_ref.open(f)) for f in zip_ref.namelist() if f.lower().endswith('.pdf')]
        elif pdf_file:
            pdf_sources = [(pdf_file.name, pdf_file)]

        total_files = len(pdf_sources)
        summary_list = []
        missing_products_list = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, (filename, pdf_stream) in enumerate(pdf_sources, start=1):
            status_text.text(f"Processing {filename} ({i}/{total_files})...")
            invoice_df, invoice_number, pic_number, freight, gst, total_tax, order_number, ship_to, skip = extract_invoice_data(pdf_stream)
            base_name = os.path.splitext(os.path.basename(filename))[0]

            # Save raw extracted data
            raw_invoice_path = os.path.join(temp_folder, f"{base_name}_raw_invoice.xlsx")
            invoice_df.to_excel(raw_invoice_path, index=False)

            if skip:
                summary_list.append({
                    "Invoice File": base_name,
                    "Invoice Number": invoice_number,
                    "PIC Number": pic_number,
                    "Order Number": order_number,
                    "Ship To": ship_to,
                    "Products in Invoice": 0,
                    "Products Matched to PAF": 0,
                    "Total Tax Included (Original)": total_tax,
                    "Calculated Total Payable": 0
                })
                continue

            merged_df = pd.merge(invoice_df, paf_df, left_on="Product", right_on="Valiant/RGR SKU", how="left")
            merged_df["Units Per Case"] = pd.to_numeric(merged_df["Units Per Case"], errors="coerce")
            merged_df["Quantity"] = pd.to_numeric(merged_df["Quantity"], errors="coerce")
            merged_df["Total Quantity"] = merged_df["Quantity"] * merged_df["Units Per Case"]
            merged_df["Unit Cost Price"] = merged_df["Gross Price"] / merged_df["Units Per Case"]

            final_df = merged_df[["GlobalTill SKU", "Total Quantity", "Unit Cost Price"]].copy()
            final_df.rename(columns={
                "GlobalTill SKU": "SKU",
                "Total Quantity": "Total Quantity",
                "Unit Cost Price": "Unit Cost Price"
            }, inplace=True)

            wb = load_workbook("uploaded_template.xlsx")
            ws = wb.active

            ws["B6"].value = freight
            ws["B7"].value = "rgr canada"
            ws["B8"].value = f"{invoice_number}/{pic_number}"
            ws["B9"].value = 0
            ws["B10"].value = gst
            ws["B11"].value = 0

            start_row = 14
            for idx, product in final_df.iterrows():
                ws[f"A{start_row + idx}"] = product["SKU"]
                ws[f"B{start_row + idx}"] = product["Total Quantity"]
                ws[f"C{start_row + idx}"] = product["Unit Cost Price"]

            final_invoice_path = os.path.join(final_invoice_folder, f"{base_name}_final_invoice.xlsx")
            wb.save(final_invoice_path)

            product_count_invoice = len(invoice_df)
            product_count_processed = final_df["SKU"].notna().sum()

            sumproduct = (merged_df["Total Quantity"] * merged_df["Unit Cost Price"]).sum()
            calculated_total = sumproduct + gst + freight

            summary_list.append({
                "Invoice File": base_name,
                "Invoice Number": invoice_number,
                "PIC Number": pic_number,
                "Order Number": order_number,
                "Ship To": ship_to,
                "Products in Invoice": product_count_invoice,
                "Products Matched to PAF": product_count_processed,
                "Total Tax Included (Original)": total_tax,
                "Calculated Total Payable": round(calculated_total, 2)
            })

            # Add missing products
            missing_products = merged_df[merged_df["GlobalTill SKU"].isna()]
            for idx, row in missing_products.iterrows():
                missing_products_list.append({
                    "Invoice File": base_name,
                    "Product": row["Product"],
                    "Description": row["Description"],
                    "Quantity": row["Quantity"],
                    "Gross Price": row["Gross Price"]
                })

            progress_bar.progress(i / total_files)

        # Create the Excel summary file
        summary_df = pd.DataFrame(summary_list)
        missing_products_df = pd.DataFrame(missing_products_list)
        paf_data_df = paf_df

        with pd.ExcelWriter(summary_file, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Summary Report", index=False)
            missing_products_df.to_excel(writer, sheet_name="Missing Products", index=False)
            paf_data_df.to_excel(writer, sheet_name="PAF Data", index=False)

        st.success("✅ Processing Complete!")

        # Generate a download link for the Summary Report
        with open(summary_file, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="invoice_summary.xlsx">📥 Download Invoice Summary</a>'
            st.markdown(href, unsafe_allow_html=True)

        st.write("### Download Final Invoices")
        for fname in os.listdir(final_invoice_folder):
            if fname.endswith(".xlsx"):
                path = os.path.join(final_invoice_folder, fname)
                with open(path, "rb") as f:
                    st.download_button(f"Download {fname}", f, fname)

        # Final Invoices ZIP
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, "w") as zipf:
            for fname in os.listdir(final_invoice_folder):
                if fname.endswith(".xlsx"):
                    zipf.write(os.path.join(final_invoice_folder, fname), fname)

        zip_buffer.seek(0)
        st.download_button("Download All Final Invoices as ZIP", zip_buffer, "final_invoices.zip")
