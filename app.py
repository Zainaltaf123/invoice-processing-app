import streamlit as st
import pandas as pd
import os
import zipfile
import shutil
from io import BytesIO
import re

st.set_page_config(layout="wide")

st.title("Valiant PAF Invoice Validation Tool")

# --- File Upload ---
paf_file = st.file_uploader("Upload PAF File", type=["xlsx"])
invoice_files = st.file_uploader("Upload Invoice Excel Files", accept_multiple_files=True, type=["xlsx"])
template_file = st.file_uploader("Upload Invoice Template File", type=["xlsx"])

if st.button("Process Files"):
    if not (paf_file and invoice_files and template_file):
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

        mismatch_dfs = {}  # For mismatch tabs

        for invoice_file in invoice_files:
            invoice_filename = invoice_file.name

            # Load invoice
            invoice_df = pd.read_excel(invoice_file)

            try:
                header_info = invoice_df.iloc[0:7].copy()
                invoice_number = str(header_info.iloc[0, 1])
                pic_number = str(header_info.iloc[1, 1])
                order_number = str(header_info.iloc[2, 1])
                ship_to = str(header_info.iloc[3, 1])
                freight = float(header_info.iloc[5, 1])
                tax = float(header_info.iloc[6, 1])
            except Exception as e:
                st.error(f"Header extraction error in {invoice_filename}: {e}")
                continue

            # Extract product data from row 10 onwards
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

            # Add mismatch tab logic
            if len(product_data) != len(matched_df):
                mismatch_dfs[f"Mismatch_{invoice_filename}"] = matched_df

            # --- Create Final Invoice Excel ---
            template_xl = pd.read_excel(template_file, sheet_name=None)
            new_invoice = template_xl.copy()

            for sheet in new_invoice:
                df = new_invoice[sheet]
                df.replace("{{Invoice Number}}", invoice_number, inplace=True)
                df.replace("{{PIC Number}}", pic_number, inplace=True)
                df.replace("{{Order Number}}", order_number, inplace=True)
                df.replace("{{Ship To}}", ship_to, inplace=True)
                df.replace("{{Freight}}", freight, inplace=True)
                df.replace("{{Tax}}", tax, inplace=True)

            # Insert matched_df into the right place (assuming there's a sheet named "Products")
            if "Products" in new_invoice:
                new_invoice["Products"] = matched_df

            output_path = f"final_invoice_output/{os.path.splitext(invoice_filename)[0]}_final_invoice.xlsx"
            with pd.ExcelWriter(output_path) as writer:
                for sheet, df in new_invoice.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)

        # --- Create Summary Report ---
        summary_df = pd.DataFrame(summary_data)
        missing_df = pd.DataFrame(missing_products_all)

        summary_path = "invoice_summary.xlsx"
        with pd.ExcelWriter(summary_path) as writer:
            summary_df.to_excel(writer, sheet_name="Summary Report", index=False)
            missing_df.to_excel(writer, sheet_name="Missing Products", index=False)
            paf_df.to_excel(writer, sheet_name="PAF Data", index=False)

            # Add mismatch sheets
            for sheet_name, mismatch_df in mismatch_dfs.items():
                mismatch_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # Excel max sheet name = 31 chars

        # --- Download Link ---
        with open(summary_path, "rb") as f:
            st.download_button("Download Invoice Summary", data=f, file_name="invoice_summary.xlsx")

        # Optional: zip final invoices for download
        zip_filename = "final_invoice_output.zip"
        with zipfile.ZipFile(zip_filename, "w") as zipf:
            for file in os.listdir("final_invoice_output"):
                zipf.write(os.path.join("final_invoice_output", file), arcname=file)

        with open(zip_filename, "rb") as f:
            st.download_button("Download All Final Invoices (ZIP)", data=f, file_name=zip_filename)
