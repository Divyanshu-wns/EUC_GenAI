import streamlit as st
import pandas as pd
import openai
from io import StringIO, BytesIO
import re

# ========== Configuration ==========
st.set_page_config(page_title="Bank Reconciliation AI", layout="centered")

# ---
# How to set your OpenAI API key securely:
# - For local development: put your key in .streamlit/secrets.toml like:
#     [openai]
#     api_key = "sk-..."
# - For Streamlit Cloud deployment: add OPENAI_API_KEY to the app's Secrets UI as:
#     OPENAI_API_KEY = "sk-..."
#   and access it as st.secrets["OPENAI_API_KEY"]
#
# The code below will work for both local and cloud deployments.
# ---

# Try both keys for compatibility (local: st.secrets["openai"]["api_key"], cloud: st.secrets["OPENAI_API_KEY"])
def get_openai_api_key():
    if "openai" in st.secrets and "api_key" in st.secrets["openai"]:
        return st.secrets["openai"]["api_key"]
    elif "OPENAI_API_KEY" in st.secrets:
        return st.secrets["OPENAI_API_KEY"]
    else:
        st.error("OpenAI API key not found in Streamlit secrets. Please set it in .streamlit/secrets.toml (local) or in the Streamlit Cloud Secrets UI (deployment).")
        st.stop()

openai.api_key = get_openai_api_key()

# ========== UI Header ==========
st.title("📊 Bank Reconciliation using OpenAI")
st.markdown("Upload an Excel file with **EDW** and **Journal** sheets. The app will use GPT to match transactions.")

# ========== File Upload ==========
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    logs = []

    # Step 1: Load and Convert Sheets
    logs.append("✅ File uploaded successfully.")
    logs.append("🔄 Reading EDW and Journal sheets...")

    try:
        edw_df = pd.read_excel(uploaded_file, sheet_name="EDW")
        journal_df = pd.read_excel(uploaded_file, sheet_name="Journal")
    except Exception as e:
        st.error("❌ Could not read EDW or Journal sheets.")
        st.stop()

    edw_csv = edw_df.to_csv(index=False)
    journal_csv = journal_df.to_csv(index=False)

    logs.append("📄 Sheets converted to CSV.")

    # Step 2: Prepare GPT Prompt
    prompt = f"""
You are a finance assistant that performs automated reconciliation.

Match entries from the Journal sheet with combinations of transactions from the EDW sheet using these rules:

- Match on Account Number, Tran Code, and Journal Date = Process Date
- Match absolute values of EDW Amounts that sum up to Journal Debit Amount
- For each match:
  - Create one **CR row** for the Journal entry:
    - Positive amount
    - Ref 3 = Journal Description
    - Ref 4 = GL Code
  - Create one or more **DR rows** for the matched EDW transactions:
    - Use **original negative sign** from EDW "Amount (INR)"
    - Ref 1 = EDW Transaction ID
    - Ref 2 = Ref Code
    - Ref 3 = EDW Description

Make sure that:
- Amount in DR lines is **negative**
- Amount in CR lines is **positive**
- All amounts are in INR
- Bus Entity = India_BU
- Rule = AutoRule <Tran Code>
- Status = Matched (if matched else UNMATCHED)

Output the final reconciliation as a **CSV** with columns:

No,Item Type,Reconciliation,SIDE,Value Date,Ref 1,Amount,Amt CCY,Bus Entity,Stmt Date,Rule,ENTRY DATE,Ref 2,Ref 3,Ref 4,Tran Code,Status

### EDW Sheet CSV:
{edw_csv}

### Journal Sheet CSV:
{journal_csv}
"""

    logs.append("🤖 Sending data to OpenAI GPT...")

    with st.spinner("Calling GPT... Please wait."):
        try:
            client = openai.OpenAI(api_key=st.secrets["openai"]["api_key"])
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            gpt_output = response.choices[0].message.content
        except Exception as e:
            st.error("❌ Error calling OpenAI API.")
            st.exception(e)
            st.stop()

    logs.append("📩 Response received from GPT.")

    # Extract CSV from GPT output
    csv_pattern = r"```csv\s*(.*?)```"
    match = re.search(csv_pattern, gpt_output, re.DOTALL)

    if match:
        csv_text = match.group(1).strip()
    else:
        # Fallback: find last table-like block with "No," at start
        lines = gpt_output.strip().splitlines()
        csv_start = None
        for i, line in enumerate(lines):
            if line.strip().startswith("No,"):
                csv_start = i
                break
        if csv_start is not None:
            csv_text = "\n".join(lines[csv_start:])
        else:
            st.error("❌ Could not find CSV data in GPT response.")
            st.code(gpt_output)
            st.stop()

    try:
        recon_df = pd.read_csv(StringIO(csv_text))
        logs.append("📊 GPT response parsed to DataFrame.")
    except Exception as e:
        st.error("❌ CSV parsing failed.")
        st.code(csv_text)
        st.exception(e)
        st.stop()

    logs.append("💾 Generating output Excel file...")
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        recon_df.to_excel(writer, sheet_name='Reconciliation', index=False)
        edw_df.to_excel(writer, sheet_name='EDW', index=False)
        journal_df.to_excel(writer, sheet_name='Journal', index=False)
    excel_buffer.seek(0)

    # Step 4: Display Logs and Download
    with st.expander("📝 Processing Logs", expanded=True):
        for log in logs:
            st.write(log)

    st.download_button(
        label="📥 Download Reconciled Excel",
        data=excel_buffer,
        file_name="Recon_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
