import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import re

# ------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------
st.set_page_config(page_title="Scan to Excel + Subtotals", layout="wide")
st.title("📸 Scanned PDF to Excel with Subtotals")

# ------------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------------

def process_scan_to_text(file_bytes):
    """ Converts PDF to text using OCR """
    try:
        images = convert_from_bytes(file_bytes)
    except Exception as e:
        st.error(f"Error converting PDF to image: {e}")
        return ""

    full_text = ""
    progress_bar = st.progress(0)
    
    for i, image in enumerate(images):
        # psm 6 = Assume a single uniform block of text
        text = pytesseract.image_to_string(image, config='--psm 6')
        full_text += text + "\n"
        progress_bar.progress((i + 1) / len(images))
    
    progress_bar.empty()
    return full_text

def text_to_dataframe(text, split_strategy="2+ Spaces"):
    """
    Parses raw text into DataFrame based on user selected strategy.
    """
    lines = text.split('\n')
    data = []
    
    # Define the regex pattern based on user choice
    if split_strategy == "2+ Spaces":
        pattern = r'\s{2,}'  # Standard for tables
    elif split_strategy == "1+ Space":
        pattern = r'\s+'     # Aggressive splitting
    else:
        pattern = r'\t+'     # Tab separated
    
    for line in lines:
        if line.strip():
            # Remove leading/trailing whitespace before splitting
            clean_line = line.strip()
            row = re.split(pattern, clean_line)
            data.append(row)

    if not data:
        return None

    # Normalize row lengths (pad missing columns with empty strings)
    max_cols = max(len(row) for row in data)
    normalized_data = [row + [''] * (max_cols - len(row)) for row in data]
    
    # Create generic headers (Column 1, Column 2...) to avoid weird "4W" names
    headers = [f"Col_{i+1}" for i in range(max_cols)]
    
    # We use the first row as data, not header, because OCR often messes up headers.
    # The user can rename them in the editor.
    df = pd.DataFrame(normalized_data, columns=headers)
    return df

# ------------------------------------------------------------------
# MAIN APP LOGIC
# ------------------------------------------------------------------

uploaded_file = st.file_uploader("Upload Scanned PDF", type=["pdf"])

# STORE DATA IN SESSION STATE SO IT DOESN'T RESET WHEN YOU CHANGE SETTINGS
if 'ocr_text' not in st.session_state:
    st.session_state.ocr_text = ""

if uploaded_file is not None:
    # Only run OCR if it's a new file
    if st.session_state.ocr_text == "":
        st.info("Reading scanned image... this may take a moment.")
        file_bytes = uploaded_file.read()
        st.session_state.ocr_text = process_scan_to_text(file_bytes)

    raw_text = st.session_state.ocr_text
    
    if raw_text:
        # -------------------------------------------------------
        # SECTION 1: TEXT PARSING SETTINGS
        # -------------------------------------------------------
        with st.expander("⚙️ Parsing Settings (Click here if columns are messy)", expanded=True):
            split_choice = st.radio(
                "How should we split the columns?",
                ["2+ Spaces", "1+ Space"],
                index=0,
                horizontal=True,
                help="If your data is stuck in one column, try '1+ Space'."
            )
        
        # Parse text into DataFrame
        df = text_to_dataframe(raw_text, split_strategy=split_choice)
        
        if df is not None:
            st.success("Scan processed!")

            # -------------------------------------------------------
            # SECTION 2: CONFIGURE COLUMNS
            # -------------------------------------------------------
            st.subheader("1. Configure Columns")
            
            c1, c2 = st.columns(2)
            with c1:
                cat_col = st.selectbox("Category Column (Group By):", ["Select..."] + list(df.columns))
            with c2:
                val_col = st.selectbox("Value Column (to Sum):", ["Select..."] + list(df.columns))

            # -------------------------------------------------------
            # SECTION 3: VERIFY & EDIT
            # -------------------------------------------------------
            st.subheader("2. Verify Data")
            st.caption("Rename columns by double-clicking the headers below.")
            
            # Simple editor - let user fix data before calc
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

            # -------------------------------------------------------
            # SECTION 4: RESULTS
            # -------------------------------------------------------
            st.divider()
            
            if cat_col != "Select..." and val_col != "Select...":
                if cat_col == val_col:
                    st.warning("⚠️ You selected the same column for both Category and Value. Please select different columns.")
                else:
                    try:
                        # 1. Clean the numbers
                        # Remove ' . ' or other OCR noise from numbers
                        edited_df[val_col] = (
                            edited_df[val_col]
                            .astype(str)
                            .str.replace(r'[^\d\.\-]', '', regex=True) # Keep only digits, dots, and minus
                        )
                        edited_df[val_col] = pd.to_numeric(edited_df[val_col], errors='coerce').fillna(0)

                        # 2. Calculate safely
                        # We explicitly name the series to avoid the "already exists" crash
                        subtotal_series = edited_df.groupby(cat_col)[val_col].sum()
                        subtotal_series.name = "Total Amount" 
                        
                        result_df = subtotal_series.reset_index()

                        # 3. Show Result
                        st.subheader("3. Calculation Results")
                        st.dataframe(result_df, use_container_width=True)
                        
                        # 4. Download
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            edited_df.to_excel(writer, sheet_name='Clean Data', index=False)
                            result_df.to_excel(writer, sheet_name='Subtotals', index=False)
                            
                        st.download_button("📥 Download Excel", output.getvalue(), "scan_results.xlsx")

                    except Exception as e:
                        st.error(f"Calculation Error: {e}")
    
    # Reset button to clear cache for new file
    if st.button("Clear & Upload New File"):
        st.session_state.ocr_text = ""
        st.rerun()
