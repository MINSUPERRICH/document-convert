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

# NOTE: If running locally (not on Cloud), you might need to point 
# to your tesseract installation path manually, e.g.:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ------------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------------

def process_scan_to_text(file_bytes):
    """
    Converts PDF bytes to images, then runs OCR on the images.
    Returns a single string of all text found.
    """
    # 1. Convert PDF to a list of images (one per page)
    try:
        images = convert_from_bytes(file_bytes)
    except Exception as e:
        st.error(f"Error converting PDF to image. Is Poppler installed? Error: {e}")
        return ""

    full_text = ""
    
    # 2. Run Tesseract OCR on each page image
    # progress bar for multi-page docs
    progress_bar = st.progress(0)
    for i, image in enumerate(images):
        # psm 6 = Assume a single uniform block of text (good for tables)
        text = pytesseract.image_to_string(image, config='--psm 6')
        full_text += text + "\n"
        progress_bar.progress((i + 1) / len(images))
    
    progress_bar.empty()
    return full_text

def text_to_dataframe(text):
    """
    Parses raw OCR text into a DataFrame.
    Assumes columns are separated by multiple spaces.
    """
    lines = text.split('\n')
    data = []
    
    for line in lines:
        if line.strip():  # Skip empty lines
            # Split by 2 or more spaces to find columns
            row = re.split(r'\s{2,}', line.strip())
            data.append(row)

    if not data:
        return None

    # Determine max columns to normalize row lengths
    max_cols = max(len(row) for row in data)
    
    # Pad shorter rows with None so pandas doesn't crash
    normalized_data = [row + [None] * (max_cols - len(row)) for row in data]
    
    # Assume first row is header
    df = pd.DataFrame(normalized_data[1:], columns=normalized_data[0])
    return df

# ------------------------------------------------------------------
# MAIN APP LOGIC
# ------------------------------------------------------------------

uploaded_file = st.file_uploader("Upload Scanned PDF", type=["pdf"])

if uploaded_file is not None:
    st.info("Reading scanned image... this may take a moment.")
    
    # Read file bytes
    file_bytes = uploaded_file.read()
    
    # 1. OCR Extraction
    raw_text = process_scan_to_text(file_bytes)
    
    if raw_text:
        # 2. Convert to Table
        df = text_to_dataframe(raw_text)
        
        if df is not None:
            st.success("Scan processed!")
            
            st.subheader("1. Verify & Edit Data")
            st.caption("OCR can be messy. Please correct column headers and values below before calculating.")

            # Try to convert numbers automatically
            for col in df.columns:
                # Remove currency symbols or commas for conversion
                clean_col = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
                df[col] = pd.to_numeric(clean_col, errors='ignore')

            # The Editable Grid
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

            st.divider()

            # 3. Calculation Logic
            st.subheader("2. Calculate Subtotals")

            c1, c2 = st.columns(2)
            with c1:
                group_col = st.selectbox("Group By (Category):", options=["Select..."] + list(edited_df.columns))
            
            with c2:
                # Filter for numeric columns only for the sum operation
                numeric_cols = edited_df.select_dtypes(include=['float64', 'int64']).columns.tolist()
                sum_col = st.selectbox("Calculate Sum For (Value):", options=["Select..."] + numeric_cols)

            if group_col != "Select..." and sum_col != "Select...":
                try:
                    # Calculation
                    subtotal_df = edited_df.groupby(group_col)[sum_col].sum().reset_index()
                    
                    # Formatting
                    subtotal_df.rename(columns={sum_col: f"Total {sum_col}"}, inplace=True)
                    
                    # Show results
                    st.write("### Subtotal Results")
                    st.dataframe(subtotal_df, use_container_width=True)

                    # 4. Export
                    st.divider()
                    st.subheader("3. Download Results")
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        edited_df.to_excel(writer, sheet_name='Cleaned Data', index=False)
                        subtotal_df.to_excel(writer, sheet_name='Subtotals', index=False)
                        
                    st.download_button(
                        label="📥 Download Excel File",
                        data=output.getvalue(),
                        file_name="scanned_data_subtotals.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Calculation Error: {e}. Please check if the 'Value' column contains non-numbers.")
            
        else:
            st.error("OCR found text, but couldn't identify a table structure. Try a clearer scan.")
            with st.expander("See raw text detected"):
                st.text(raw_text)
    else:
        st.error("No text detected. The image might be too blurry or blank.")
