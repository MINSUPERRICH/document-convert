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

# OPTIONAL: If running locally on Windows, uncomment and set your path:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ------------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------------

def process_scan_to_text(file_bytes):
    """
    Converts PDF bytes to images, then runs OCR on the images.
    Returns a single string of all text found.
    """
    try:
        images = convert_from_bytes(file_bytes)
    except Exception as e:
        st.error(f"Error converting PDF to image. Is Poppler installed? Error: {e}")
        return ""

    full_text = ""
    # Progress bar for multi-page docs
    progress_bar = st.progress(0)
    
    for i, image in enumerate(images):
        # --psm 6 assumes a single uniform block of text (good for tables)
        text = pytesseract.image_to_string(image, config='--psm 6')
        full_text += text + "\n"
        progress_bar.progress((i + 1) / len(images))
    
    progress_bar.empty()
    return full_text

def text_to_dataframe(text):
    """
    Parses raw OCR text into a DataFrame.
    Assumes columns are separated by multiple spaces (2+).
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

    # Normalize row lengths
    max_cols = max(len(row) for row in data)
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
    
    file_bytes = uploaded_file.read()
    raw_text = process_scan_to_text(file_bytes)
    
    if raw_text:
        df = text_to_dataframe(raw_text)
        
        if df is not None:
            st.success("Scan processed successfully!")
            st.divider()

            # -------------------------------------------------------
            # SECTION 1: CONFIGURE COLUMNS
            # -------------------------------------------------------
            st.subheader("1. Configure Columns")
            st.markdown("Identify which columns represent your **Categories** and your **Values** so we can format the editor correctly.")
            
            col_conf1, col_conf2 = st.columns(2)
            
            with col_conf1:
                # User identifies the text column (e.g., Description/Department)
                cat_col_name = st.selectbox(
                    "Which column is the Category (Group By)?", 
                    options=["Select..."] + list(df.columns),
                    index=0
                )

            with col_conf2:
                # User identifies the number column (e.g., Amount/Price)
                num_col_name = st.selectbox(
                    "Which column is the Value (to Sum)?", 
                    options=["Select..."] + list(df.columns),
                    index=0
                )

            # -------------------------------------------------------
            # SECTION 2: EDIT DATA (The Advanced Editor)
            # -------------------------------------------------------
            st.subheader("2. Verify & Edit Data")
            st.caption("Double-click cells to fix OCR errors. Rows with empty values will be ignored in calculations.")

            # Create the Column Configuration Dictionary
            column_settings = {}

            # A. Configure Number Column (Force input to be numbers)
            if num_col_name != "Select...":
                # Pre-clean the data: Remove $ and , so it converts to float
                clean_col = df[num_col_name].astype(str).str.replace(r'[$,]', '', regex=True)
                df[num_col_name] = pd.to_numeric(clean_col, errors='ignore')

                column_settings[num_col_name] = st.column_config.NumberColumn(
                    label=f"{num_col_name} (Number)",
                    help="Only numbers allowed here.",
                    min_value=0,
                    format="$%.2f"
                )

            # B. Configure Category Column (Dropdown for consistency)
            if cat_col_name != "Select...":
                # Get unique values for the dropdown list
                unique_cats = df[cat_col_name].dropna().unique().tolist()
                
                column_settings[cat_col_name] = st.column_config.SelectboxColumn(
                    label=f"{cat_col_name} (Category)",
                    help="Select a valid category",
                    options=unique_cats,
                    width="medium",
                    required=True
                )

            # Render the Editor
            edited_df = st.data_editor(
                df,
                column_config=column_settings,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True
            )

            # -------------------------------------------------------
            # SECTION 3: CALCULATE & EXPORT
            # -------------------------------------------------------
            st.divider()
            st.subheader("3. Results & Download")

            # Only calculate if columns are selected
            if cat_col_name != "Select..." and num_col_name != "Select...":
                try:
                    # 1. Group By Calculation
                    # Convert column to numeric one last time to be safe before summing
                    edited_df[num_col_name] = pd.to_numeric(edited_df[num_col_name], errors='coerce').fillna(0)
                    
                    subtotal_df = edited_df.groupby(cat_col_name)[num_col_name].sum().reset_index()
                    subtotal_df.columns = [cat_col_name, f"Total {num_col_name}"] # Rename for display

                    # 2. Show Table
                    st.write("### Subtotals")
                    st.dataframe(subtotal_df, use_container_width=True)

                    # 3. Download Button
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
                    st.error(f"Could not calculate: {e}. Check if the Value column contains non-numbers.")
            else:
                st.warning("Please select both a Category and a Value column in Section 1 to see subtotals.")
        else:
            st.error("OCR finished but could not find a table structure. The scan might be too blurry.")
            with st.expander("View Raw Text"):
                st.text(raw_text)
