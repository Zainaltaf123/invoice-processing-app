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

st.title("Invoice Processing Tool with Bulk ZIP Download")

zip_file = st.file_uploader("Upload Invoices ZIP", type=["zip"])
paf_file = st.file_uploader("Upload PAF Excel", type=["xlsx"])
template_file = st.file_uploader("Upload Invoice Template Excel", type=["xlsx"])

temp_folder = "temp_excels"
final_folder = "final_outputs"
final_invoice_folder = "final_invoice_output"
summary_file = "Summary_Report.xlsx"

for folder in [temp_folder, final_folder, final_invoice_folder]:
    os.makedirs(folder, exist_ok=True)

line_pattern = re.compile(
    r"^\s*(\d+)\s+"
    r"(\S+)\s+"
    r"(.+?)\s+"
    r"(\d+(?:\.?\d*)?(?:EA|PAC))\s+"
    r"([\d\.]+)\s+"
    r"([\d\.]+)\s*$"
)

def extract_invoice_data(file):
    products = []
    invoice_number = pic_number = freight_charges = gst_amount = total_tax_included = order_number = ship_to_address = None
    skip = False

    with pdfplumber.open(file) as pdf:
        first_text = pdf.pages[0].extract_text()

        inv_m = re.search(r"(INV\d{6})", first_text)
        pic_m = re.search(r"(PIC\d{6})", first_text)
        invoice_number = inv_m.group(1) if inv_m else None
        pic_number = pic_m.group(1) if pic_m else None

        order_number_lines = first_text.splitlines()
        order_number = ""
        for i, line in enumerate(order_number_lines):
            if "Order Number" in line and i + 1 < len(order_number_lines):
                next_line = order_number_lines[i + 1].strip()
                if next_line.startswith("RGRHO") or next_line.startswith("CCAO"):
                    order_number = next_line
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

if zip_file and paf_file and template_file:
    if st.button("Process Invoices"):

        for folder in [temp_folder, final_folder, final_invoice_folder]:
            for f in os.listdir(folder):
                os.remove(os.path.join(folder, f))

        with open("uploaded_invoices.zip", "wb") as f:
            f.write(zip_file.getbuffer())
        with open("uploaded_paf.xlsx", "wb") as f:
            f.write(paf_file.getbuffer())
        with open("uploaded_template.xlsx", "wb") as f:
            f.write(template_file.getbuffer())

        paf_df = pd.read_excel("uploaded_paf.xlsx")
        paf_df.columns = paf_df.columns.str.strip()
        paf_df = paf_df.drop_duplicates(subset=["Valiant/RGR SKU"])

        zip_path = "uploaded_invoices.zip"
        template_path = "uploaded_template.xlsx"

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            pdf_files = [f for f in zip_ref.namelist() if f.lower().endswith('.pdf')]
            total_files = len(pdf_files)

            summary_list = []

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, pdf_file in enumerate(pdf_files, start=1):
                status_text.text(f"Processing {pdf_file} ({i}/{total_files})...")

                with zip_ref.open(pdf_file) as file:
                    invoice_df, invoice_number, pic_number, freight, gst, total_tax, order_number, ship_to, skip = extract_invoice_data(file)

                base_name = os.path.splitext(os.path.basename(pdf_file))[0]

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

                wb = load_workbook(template_path)
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
                    ws[f"C{start_row + idx}"] = round(float(product["Unit Cost Price"]), 2)

                final_invoice_path = os.path.join(final_invoice_folder, f"{base_name}_final_invoice.xlsx")
                wb.save(final_invoice_path)

                product_count_invoice = len(invoice_df)
                product_count_processed = merged_df["GlobalTill SKU"].notna().sum()

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

                progress_bar.progress(i / total_files)

            summary_df = pd.DataFrame(summary_list)
            summary_df.to_excel(summary_file, index=False)

            st.success("✅ Processing Complete!")

            with open(summary_file, "rb") as f:
                st.download_button("Download Summary Report", f, "Summary_Report.xlsx")

            st.write("### Download Final Invoices (individually or all together)")

            for fname in os.listdir(final_invoice_folder):
                if fname.endswith(".xlsx"):
                    path = os.path.join(final_invoice_folder, fname)
                    with open(path, "rb") as f:
                        st.download_button(f"Download {fname}", f, fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            zip_buffer = io.BytesIO()
            with ZipFile(zip_buffer, "w") as zipf:
                for fname in os.listdir(final_invoice_folder):
                    zipf.write(os.path.join(final_invoice_folder, fname), arcname=fname)
            zip_buffer.seek(0)

            st.download_button(
                label="Download All Final Invoices (ZIP)",
                data=zip_buffer,
                file_name="Final_Invoices.zip",
                mime="application/zip"
            )
