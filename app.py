import streamlit as st
import pandas as pd
import os
import zipfile
from io import BytesIO
import shutil
import PyPDF2

st.set_page_config(layout="wide")

st.title("Valiant PAF Invoice Validation Tool")

# --- File Upload ---
paf_file = st.file_uploader("Upload PAF File", type=["xlsx"])
invoice_file = st.file_uploader("Upload Invoice Files (PDF or ZIP)", accept_multiple_files=True, type=["pdf", "zip"])
template_file = st.file_uploader("Upload Invoice Template File", type=["xlsx"])

if st.button("Process Files"):
    if not (paf_file and invoice_file and template_file):
        st.error("Please upload all files (PAF, Invoices, Template).")
    else:
        st.success("Processing...")

        # Create necessary folders
        os.makedirs("temp_excels", exist_ok=True)
        os.makedirs("final_outputs", exist_ok=True)
        os.makedirs("final_invoice_output", exist_ok=True)

        # Load PAF File
        paf_df = pd.read_excel(paf_file)
        paf_df = paf_df.drop_duplicates(subset=["Valiant/RGR SKU"], keep="first")

        summary_data = []
        missing_products_all = []

        # For mismatch tabs in summary
        mismatch_invoice_dfs = {}

        # Process each uploaded file
        for uploaded_file in invoice_file:
            invoice_filename = uploaded_file.name

            # Check if file is a ZIP
            if invoice_filename.endswith('.zip'):
                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                    zip_ref.extractall('temp_excels')
                # Find extracted PDF files
                pdf_files = [f for f in os.listdir('temp_excels') if f.endswith('.pdf')]
                for pdf_file in pdf_files:
                    with open(f'temp_excels/{pdf_file}', 'rb') as f:
                        pdf_data = f.read()
                        # Process this PDF file as if it was uploaded directly
                        invoice_df = process_pdf(pdf_data)

                        # Now extract data as needed from the PDF invoice
                        process_invoice_data(invoice_df, invoice_filename, paf_df, summary_data, missing_products_all, mismatch_invoice_dfs)
            else:
                # If it's a PDF file, process directly
                with open(uploaded_file, 'rb') as f:
                    pdf_data = f.read()
                    invoice_df = process_pdf(pdf_data)

                    # Now extract data as needed from the PDF invoice
                    process_invoice_data(invoice_df, invoice_filename, paf_df, summary_data, missing_products_all, mismatch_invoice_dfs)

        # --- Create Summary Report ---
        summary_df = pd.DataFrame(summary_data)
        missing_df = pd.DataFrame(missing_products_all)

        summary_path = "invoice_summary.xlsx"
        with pd.ExcelWriter(summary_path, engine="openpyxl", mode="w") as writer:
            # Existing outputs
            summary_df.to_excel(writer, sheet_name="Summary Report", index=False)
            missing_df.to_excel(writer, sheet_name="Missing Products", index=False)
            paf_df.to_excel(writer, sheet_name="PAF Data", index=False)

            # Additional mismatch tabs
            for invoice_filename, mismatch_df in mismatch_invoice_dfs.items():
                # Ensure safe sheet name
                sheet_name = invoice_filename.replace(".pdf", "").strip() + " — Final_Invoice"
                sheet_name = sheet_name[:31]  # Excel limit
                mismatch_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # --- Download Buttons ---
        with open(summary_path, "rb") as f:
            st.download_button("Download Invoice Summary", data=f, file_name="invoice_summary.xlsx")

        # Optional: zip final invoices for download
        zip_filename = "final_invoice_output.zip"
        with zipfile.ZipFile(zip_filename, "w") as zipf:
            for file in os.listdir("final_invoice_output"):
                zipf.write(os.path.join("final_invoice_output", file), arcname=file)

        with open(zip_filename, "rb") as f:
            st.download_button("Download All Final Invoices (ZIP)", data=f, file_name=zip_filename)

# Function to process PDF invoices
def process_pdf(pdf_data):
    # Assuming you use PyPDF2 to read the PDF data (you might need to adjust this based on your needs)
    pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_data))
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()

    # Convert text to DataFrame (you may need to adjust this)
    invoice_df = pd.DataFrame({"Text": text.splitlines()})
    return invoice_df

# Function to process each invoice file and match products
def process_invoice_data(invoice_df, invoice_filename, paf_df, summary_data, missing_products_all, mismatch_invoice_dfs):
    try:
        # Extract relevant data from invoice_df (implement your extraction logic)
        header_info = invoice_df.iloc[0:7].copy()
        invoice_number = str(header_info.iloc[0, 1])
        pic_number = str(header_info.iloc[1, 1])
        order_number = str(header_info.iloc[2, 1])
        ship_to = str(header_info.iloc[3, 1])
        freight = float(header_info.iloc[5, 1])
        tax = float(header_info.iloc[6, 1])
    except Exception as e:
        st.error(f"Header extraction error in {invoice_filename}: {e}")
        return

    # Extract product data (this part might need to be adjusted for your invoice format)
    product_data = invoice_df.iloc[9:].dropna(subset=["Product Description"])
    product_data = product_data[["Product Description", "Qty", "Gross Price"]].copy()
    product_data.columns = ["Product Description", "Qty", "Gross Price"]
    product_data["Qty"] = pd.to_numeric(product_data["Qty"], errors="coerce").fillna(0)

    # Match products to PAF
    matched = []
    missing = []

    for _, row in product_data.iterrows():
        desc = str(row["Product Description"]).strip().upper()
        qty = row["Qty"]
        gross = row["Gross Price"]

        matched_row = paf_df[paf_df["Product Description"].str.upper().str.strip() == desc]
        if not matched_row.empty:
            sku = matched_row["Valiant/RGR SKU"].values[0]
            unit_cost = matched_row["Unit Cost"].values[0]
            matched.append({
                "Valiant/RGR SKU": sku,
                "Quantity": qty,
                "Unit Cost": unit_cost,
                "Product Description": desc
            })
        else:
            missing.append({
                "Invoice File": invoice_filename,
                "Product Code": "",
                "Description": desc,
                "Quantity": qty,
                "Gross Price": gross
            })

    matched_df = pd.DataFrame(matched)
    recalculated_total = (matched_df["Quantity"] * matched_df["Unit Cost"]).sum() + freight + tax
    original_total = (product_data["Qty"] * product_data["Gross Price"]).sum() + freight + tax

    summary_data.append({
        "Invoice File": invoice_filename,
        "Invoice Number": invoice_number,
        "PIC Number": pic_number,
        "Order Number": order_number,
        "Ship To": ship_to,
        "Product Count (Invoice)": len(product_data),
        "Product Count (Matched)": len(matched_df),
        "Original Total": round(original_total, 2),
        "Recalculated Total": round(recalculated_total, 2),
    })

    missing_products_all.extend(missing)

    # Save mismatched final_invoice for summary file if counts differ
    if len(product_data) != len(matched_df):
        sheet_name = f"{invoice_filename} — Final_Invoice"
        # Ensure sheet name doesn't exceed Excel's 31-character limit
        sheet_name = sheet_name[:31]
        mismatch_invoice_dfs[sheet_name] = matched_df
