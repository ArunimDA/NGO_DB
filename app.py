import streamlit as st
import docx
import pandas as pd
import io

# --- EXTRACTION ENGINE ---
def get_field_value(tables, label_name):
    """Searches for a label and grabs the value in the next cell or same cell."""
    for df in tables:
        rows, cols = df.shape
        for r in range(rows):
            for c in range(cols):
                cell_text = str(df.iloc[r, c]).strip()
                if label_name.upper() in cell_text.upper():
                    # Check if value is in same cell after ":"
                    if ":" in cell_text and len(cell_text.split(":", 1)[1].strip()) > 1:
                        return cell_text.split(":", 1)[1].strip()
                    # Check neighboring cells (skipping colon)
                    for offset in range(1, 3):
                        if c + offset < cols:
                            val = str(df.iloc[r, c + offset]).strip()
                            if val and val != ":":
                                return val
    return "N/A"

# --- WEB UI ---
st.set_page_config(page_title="NGO Data Extractor", layout="wide")
st.title("🏦 City Bank: NGO Memo Extractor")
st.info("Upload your Word (.docx) files to generate a consolidated Excel database.")

uploaded_files = st.file_uploader("Select NGO Memos", type="docx", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    # The 21 specific fields from your document header
    fields = [
        "Memo date", "Relationship", "Group", "Main Borrower", "Co-Utilizer",
        "CRG", "E&S Risk", "CIB Status", "External Rating", "Strategy",
        "Segment", "Lending Rate", "Exposure Type", "Branch", "Key Person",
        "Enhancement History", "RM", "UH", "Risk Manager", "AH", "Risk UH"
    ]

    for uploaded_file in uploaded_files:
        try:
            doc = docx.Document(uploaded_file)
            tables = [pd.DataFrame([[c.text.strip() for c in r.cells] for r in t.rows]) for t in doc.tables]

            record = {"Source File": uploaded_file.name}
            for field in fields:
                record[field] = get_field_value(tables, field)
            
            all_data.append(record)
        except Exception as e:
            st.error(f"Error in {uploaded_file.name}: {e}")

    if all_data:
        df_final = pd.DataFrame(all_data)
        st.success(f"Successfully processed {len(all_data)} files!")
        st.dataframe(df_final)

        # Download Logic
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Extracted_Data')
        
        st.download_button(
            label="📥 Download Consolidated Excel File",
            data=output.getvalue(),
            file_name="NGO_Master_Database.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )