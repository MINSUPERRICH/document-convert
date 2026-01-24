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
        # We use standard psm 6 to get raw lines
        text = pytesseract.image_to_string(image, config='--psm 6')
        full_text += text + "\n"
        progress_bar.progress((i + 1) / len(images))
    
    progress_bar.empty()
    return full_text

def raw_text_to_initial_df(text):
    """
    Just dumps text into a single column DataFrame initially.
    We will let pandas split it later.
    """
    lines = text.split('\n')
    # Filter out empty lines
    lines = [line.strip() for line in lines if line.strip()]
    return pd.DataFrame(lines, columns=["Raw_Text"])

# ------------------------------------------------------------------
# MAIN APP LOGIC
# ------------------------------------------------------------------

if 'df_main' not in st.session_state:
    st.session_state.df_main = None

uploaded_file = st.file_uploader("Upload Scanned PDF", type=["pdf"])

if uploaded_file is not None:
    
    # 1. READ FILE (Only once)
    if st.session_state.df_main is None:
        with st.spinner("Scanning document..."):
            file_bytes = uploaded_file.read()
            raw_text = process_scan_to_text(file_bytes)
            st.session_state.df_main = raw_text_to_initial_df(raw_text)

    df = st.session_state.df_main.copy()
    
    if not df.empty:
        st.divider()

        # -------------------------------------------------------
        # SECTION 1: DATA CLEANUP (THE FIX)
        # -------------------------------------------------------
        st.subheader("1. Clean Up Structure")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.markdown("""
            **Does your data look stuck in one column?** Click the button below to force-split the text by spaces.
            """)
            
            # THE MAGIC BUTTON: Splits "Col_1" into "0, 1, 2, 3..."
            if st.button("✂️ Split Text into Columns"):
                # Split by whitespace
                df_split = df.iloc[:, 0].str.split(expand=True)
                
                # Update the main dataframe in session state
                st.session_state.df_main = df_split
                st.rerun()

            if st.button("↺ Reset to Original"):
                st.session_state.df_main = None
                st.rerun()

        # Display the current state of the data
        st.write("Current Data Preview:")
        st.dataframe(df.head(5), use_container_width=True)

        # -------------------------------------------------------
        # SECTION 2: DEFINE COLUMNS
        # -------------------------------------------------------
        st.divider()
        st.subheader("2. Select Columns")
        
        # Get list of current column names (0, 1, 2, 3...)
        cols = list(df.columns)
        
        c1, c2 = st.columns(2)
        with c1:
            cat_col = st.selectbox("Category Column (Group By):", ["Select..."] + cols)
        with c2:
            val_col = st.selectbox("Value Column (to Sum):", ["Select..."] + cols)

        # -------------------------------------------------------
        # SECTION 3: EDIT & CALCULATE
        # -------------------------------------------------------
        st.divider()
        st.subheader("3. Verify & Download")
        
        # Allow editing
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        if cat_col != "Select..." and val_col != "Select...":
            if cat_col == val_col:
                st.warning("Please pick different columns for Category and Value.")
            else:
                try:
                    # CLEANUP LOGIC:
                    # 1. Convert Value column to numeric (force errors to NaN)
                    # Remove '$', ',', and spaces
                    clean_vals = edited_df[val_col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
                    edited_df[f"Clean_{val_col}"] = pd.to_numeric(clean_vals, errors='coerce').fillna(0)

                    # 2. Group By
                    summary_df = edited_df.groupby(cat_col)[f"Clean_{val_col}"].sum().reset_index()
                    
                    # 3. Rename for display
                    summary_df.columns = ["Category", "Total Amount"]

                    st.success("Calculation Successful!")
                    st.dataframe(summary_df, use_container_width=True)

                    # 4. Download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        edited_df.to_excel(writer, sheet_name='Raw Data', index=False)
                        summary_df.to_excel(writer, sheet_name='Subtotals', index=False)

                    st.download_button(
                        label="📥 Download Excel Result",
                        data=output.getvalue(),
                        file_name="converted_scan.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"Could not calculate: {e}")
