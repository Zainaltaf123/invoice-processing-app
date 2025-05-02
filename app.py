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
    r"^\s*(\d+)\s+"               # SNo
    r"(\S+)\s+"                   # Product
    r"(.+?)\s+"                   # Description
    r"(\d+(?:\.?\d*)?(?:EA|PAC))\s+"  # Quantity+Unit
    r"([\d\.]+)\s+"               # Gross Price
    r"([\d\.]+)\s*$"              # Extension Cost
)

# Map full province names to their 2‑letter codes
PROVINCE_MAP = {
    'alberta': 'AB', 'british columbia': 'BC', 'manitoba': 'MB',
    'new brunswick': 'NB', 'newfoundland and labrador': 'NL', 'nova scotia': 'NS',
    'northwest territories': 'NT', 'nunavut': 'NU', 'ontario': 'ON',
    'prince edward island': 'PE', 'quebec': 'QC', 'saskatchewan': 'SK', 'yukon': 'YT'
}
CODES = list(PROVINCE_MAP.values())

def extract_shipto_province(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()
    lines = text.splitlines()

    # 1) find “Ship To” header
    start = next((i for i,l in enumerate(lines) if "ship to" in l.lower()), None)
    if start is None:
        return None
    # 2) find next section header
    end = next(
      (j for j in range(start+1, len(lines))
       if any(h in lines[j].lower() for h in ("customer id","invoice no","bill to"))),
      len(lines)
    )
    # 3) extract block
    block = lines[start+1:end]
    block_text = " ".join(block).lower()
    # 4) find all province hits
    hits = []
    for fullname, code in PROVINCE_MAP.items():
        idx = block_text.rfind(fullname)
        if idx >= 0:
            hits.append((idx, code))
    for code in CODES:
        m = re.search(r'\b'+code.lower()+r'\b', block_text)
        if m:
            hits.append((m.start(), code))
    if not hits:
        return None
    hits.sort(key=lambda x: x[0])
    return hits[-1][1]

def extract_invoice_data(file):
    products=[]
    invoice_number=pic_number=freight_charges=gst_amount=total_tax_included=order_number=ship_to_address=province=None
    skip=False

    # extract ship_to and province
    # streamlit file-like object needs to be rewound for pdfplumber
    file.seek(0)
    province = extract_shipto_province(file)

    file.seek(0)
    with pdfplumber.open(file) as pdf:
        first_text = pdf.pages[0].extract_text() or ""

        inv_m = re.search(r"(INV\d{6})", first_text)
        pic_m = re.search(r"(PIC\d{6})", first_text)
        invoice_number = inv_m.group(1) if inv_m else None
        pic_number = pic_m.group(1) if pic_m else None

        for line in first_text.splitlines():
            m = re.search(r"\b(RGRHO\w+|CCAO\w+)\b", line)
            if m:
                order_number = m.group(1)
                break
        if order_number and order_number.startswith("CCAO"):
            skip=True

        # capture Ship To block for summary
        ship_to_lines = re.findall(r"Ship To\s*\n(.*?)\n(.*?)\n", first_text, re.DOTALL)
        ship_to_address = "\n".join(ship_to_lines[0]) if ship_to_lines else ""

        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            try:
                header_idx = next(i for i,ln in enumerate(lines) if ln.strip().startswith("SNo.") and "Extension Cost" in ln)
            except StopIteration:
                continue
            for ln in lines[header_idx+1:]:
                if ln.strip().startswith("Page"):
                    break
                m = line_pattern.match(ln)
                if m:
                    qf=m.group(4)
                    qn=re.match(r"([\d\.]+)",qf).group(1)
                    qu=re.search(r"(EA|PAC)",qf).group(1)
                    products.append({
                        "SNo":int(m.group(1)),"Product":m.group(2),
                        "Description":m.group(3).strip(),
                        "Quantity":float(qn),"Unit":qu,
                        "Gross Price":float(m.group(5)),"Extension Cost":float(m.group(6))
                    })

        last_text=pdf.pages[-1].extract_text() or ""
        f_m=re.search(r"Freight Charges\s+([\d,]+\.\d{2})", last_text)
        g_m=re.search(r"GST/HST Amount\s+([\d,]+\.\d{2})", last_text)
        t_m=re.search(r"TOTAL TAX INCLUDED\s+([\d,]+\.\d{2})", last_text, re.IGNORECASE)

        freight_charges=float(f_m.group(1).replace(",","")) if f_m else 0
        gst_amount=float(g_m.group(1).replace(",","")) if g_m else 0
        total_tax_included=float(t_m.group(1).replace(",","")) if t_m else 0

    df=pd.DataFrame(products)
    return df, invoice_number, pic_number, freight_charges, gst_amount, total_tax_included, order_number, ship_to_address, skip, province

if (zip_file or pdf_file) and paf_file and template_file:
    if st.button("Process Invoices"):
        for d in [temp_folder, final_folder, final_invoice_folder]:
            for f in os.listdir(d):
                os.remove(os.path.join(d,f))

        with open("uploaded_paf.xlsx","wb") as f: f.write(paf_file.getbuffer())
        with open("uploaded_template.xlsx","wb") as f: f.write(template_file.getbuffer())

        paf_df=pd.read_excel("uploaded_paf.xlsx")
        paf_df.columns=paf_df.columns.str.strip()
        paf_df=paf_df.drop_duplicates(subset=["Valiant/RGR SKU"])

        # gather PDFs
        pdfs=[]
        if zip_file:
            with open("invoices.zip","wb") as f: f.write(zip_file.getbuffer())
            with zipfile.ZipFile("invoices.zip") as z:
                pdfs=[(n,z.open(n)) for n in z.namelist() if n.lower().endswith(".pdf")]
        else:
            pdfs=[(pdf_file.name,pdf_file)]

        summary=[]
        missing_list=[]
        prog=st.progress(0)
        stat=st.empty()

        for i,(name,stream) in enumerate(pdfs, start=1):
            stat.text(f"Processing {name} ({i}/{len(pdfs)})...")
            df,inv,pic,fr,gst,tot,ord,ship,skip,prov=extract_invoice_data(stream)
            bn=os.path.splitext(os.path.basename(name))[0]
            df.to_excel(os.path.join(temp_folder,f"{bn}_raw.xlsx"),index=False)
            if skip:
                summary.append({
                    "Invoice File":bn,"Invoice Number":inv,"PIC Number":pic,
                    "Order Number":ord,"Ship To":ship,"Province":prov,
                    "Products in Invoice":0,"Products Matched to PAF":0,
                    "Total Tax Included (Original)":tot,"Calculated Total Payable":0
                })
                continue

            merged=pd.merge(df,paf_df,left_on="Product",right_on="Valiant/RGR SKU",how="left")
            merged["Units Per Case"]=pd.to_numeric(merged["Units Per Case"],errors="coerce")
            merged["Total Quantity"]=merged["Quantity"]*merged["Units Per Case"]
            merged["Unit Cost Price"]=merged["Gross Price"]/merged["Units Per Case"]

            final=merged[["GlobalTill SKU","Total Quantity","Unit Cost Price"]].copy()
            final.columns=["SKU","Total Quantity","Unit Cost Price"]

            wb=load_workbook("uploaded_template.xlsx"); ws=wb.active
            ws["B6"].value=fr; ws["B7"].value="rgr canada"; ws["B8"].value=f"{inv}/{pic}"; ws["B9"].value=0
            if prov=="ON": ws["B10"].value=0; ws["B11"].value=gst
            else:        ws["B10"].value=gst; ws["B11"].value=0

            sr=14
            for idx,row in final.iterrows():
                ws[f"A{sr+idx}"].value=row["SKU"]
                ws[f"B{sr+idx}"].value=row["Total Quantity"]
                ws[f"C{sr+idx}"].value=row["Unit Cost Price"]
            outp=os.path.join(final_invoice_folder,f"{bn}_final.xlsx")
            wb.save(outp)

            cnt_inv=len(df); cnt_mat=final["SKU"].notna().sum()
            calc=final["Total Quantity"].mul(final["Unit Cost Price"]).sum()+gst+fr

            summary.append({
                "Invoice File":bn,"Invoice Number":inv,"PIC Number":pic,
                "Order Number":ord,"Ship To":ship,"Province":prov,
                "Products in Invoice":cnt_inv,"Products Matched to PAF":cnt_mat,
                "Total Tax Included (Original)":tot,"Calculated Total Payable":round(calc,2)
            })

            for _,r in merged[merged["GlobalTill SKU"].isna()].iterrows():
                missing_list.append({
                    "Invoice File":bn,"Product":r["Product"],
                    "Description":r["Description"],"Quantity":r["Quantity"],
                    "Gross Price":r["Gross Price"]
                })
            prog.progress(i/len(pdfs))

        # write summary
        s_df=pd.DataFrame(summary); m_df=pd.DataFrame(missing_list); p_df=paf_df
        with pd.ExcelWriter(summary_file,engine="openpyxl") as w:
            s_df.to_excel(w,sheet_name="Summary Report",index=False)
            m_df.to_excel(w,sheet_name="Missing Products",index=False)
            p_df.to_excel(w,sheet_name="PAF Data",index=False)

        st.success("✅ Processing Complete!")
        with open(summary_file,"rb") as f:
            b64=base64.b64encode(f.read()).decode()
            st.markdown(f'<a href="data:application/octet-stream;base64,{b64}" download="invoice_summary.xlsx">📥 Download Invoice Summary</a>',unsafe_allow_html=True)

        st.write("### Download Final Invoices")
        for fn in os.listdir(final_invoice_folder):
            if fn.endswith(".xlsx"):
                with open(os.path.join(final_invoice_folder,fn),"rb") as f:
                    st.download_button(fn,f,fn)
        buf=io.BytesIO()
        with ZipFile(buf,"w") as z:
            for fn in os.listdir(final_invoice_folder):
                if fn.endswith(".xlsx"):
                    z.write(os.path.join(final_invoice_folder,fn),fn)
        buf.seek(0)
        st.download_button("Download All Final Invoices as ZIP",buf,"final_invoices.zip")
