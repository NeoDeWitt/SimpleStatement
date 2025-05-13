import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Bank Statements to Excel", layout="wide")
st.title("Bank Statements to Excel")
st.write("Upload multiple bank statement PDFs to extract transactions and compile them into a single Excel file.")

uploaded_files = st.file_uploader("Upload PDF files", accept_multiple_files=True, type=["pdf"])

if uploaded_files:
    all_transactions = []

    for uploaded_file in uploaded_files:
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            # Updated parsing logic to handle Thai Baht and dd/mm/yy format
                            parts = line.split(" - ")
                            if len(parts) == 3:
                                try:
                                    # Adjust date format to %d/%m/%y
                                    date = datetime.strptime(parts[0].strip(), "%d/%m/%y")
                                    description = parts[1].strip()
                                    # Adjust for Thai Baht symbol and possible commas
                                    amount_str = parts[2].replace("฿", "").replace(",", "").strip()
                                    amount = float(amount_str)
                                    all_transactions.append({"Date": date, "Description": description, "Amount": amount})
                                except ValueError:
                                    st.warning(f"Skipping invalid line: {line}")
        except Exception as e:
            st.error(f"Error processing file {uploaded_file.name}: {e}")

    if all_transactions:
        df = pd.DataFrame(all_transactions)
        df = df.sort_values(by="Date")

        # Create an Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Transactions")

        st.success("Data successfully compiled into Excel!")
        st.download_button(label="Download Excel file", data=output.getvalue(), file_name="transactions.xlsx")
    else:
        st.warning("No valid transactions found in the uploaded PDFs.")
