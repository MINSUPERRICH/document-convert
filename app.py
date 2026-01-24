import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import numpy as np

# ------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------
st.set_page_config(page_title="Layout Preserving OCR", layout="wide")
st.title("📄 Exact Layout Scan-to-Excel")

# ------------------------------------------------------------------
# ALGORITHM: RECONSTRUCT TABLE FROM COORDINATES
# ------------------------------------------------------------------

def process_image_to_grid(image):
    """
    1. Runs OCR to get X,Y coordinates of every word.
    2. Groups words into Rows (based on Y).
    3. Groups words into Columns (based on X).
    4. Returns a Pandas DataFrame matching the visual layout.
    """
    # Get detailed data (words, left, top, width, height)
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    
    df_words = pd.DataFrame(data)
    # Filter out empty text and low confidence noise
    df_words = df_words[df_words['text'].str.strip() != '']
    df_words['text'] = df_words['text'].astype(str)
    
    if df_words.empty:
        return pd.DataFrame()

    # --- STEP 1: DEFINE ROWS (Y-Axis) ---
    # We round the 'top' coordinate to the nearest 10 pixels to group words on the same line
    df_words['row_group'] = (df_words['top'] / 10).round().astype(int) * 10
    
    # Sort by vertical position (rows), then horizontal (reading order)
    df_words = df_words.sort_values(by=['row_group', 'left'])

    # --- STEP 2: DEFINE COLUMNS (X-Axis) ---
    # We look at where words usually start to find "Column Lines"
    # This is a simple binning strategy: Round 'left' to nearest 50 pixels
    df_words['col_group'] = (df_words['left'] / 50).round().astype(int) * 50

    # --- STEP 3: BUILD THE GRID ---
    # We create a matrix using these groups
    
    # Get unique row and col coordinates
    unique_rows = sorted(df_words['row_group'].unique())
    unique_cols = sorted(df_words['col_group'].unique())
    
    # Create an empty DataFrame with these dimensions
    grid_df = pd.DataFrame(index=unique_rows, columns=unique_cols)
    
    # Fill the grid
    for _, row in df_words.iterrows():
        r = row['row_group']
        c = row['col_group']
        txt = row['text']
        
        # If cell is empty, add text. If not, append it (for multi-word cells)
        if pd.isna(grid_df.at[r, c]):
            grid_df.at[r, c] = txt
        else:
            grid_df.at[r, c] = f"{grid_df.at[r, c]} {txt}"

    # Reset index to look like a normal table
    grid_df = grid_df.reset_index(drop=True)
    grid_df.columns = [f"Col_{i+1}" for i in range(len(grid_df.columns))]
    
    return grid_df

# ------------------------------------------------------------------
# MAIN APP
# ------------------------------------------------------------------

if 'master_df' not in st.session_state:
    st.session_state.master_df = None

uploaded_file = st.file_uploader("Upload Scanned PDF", type=["pdf"])

if uploaded_file:
    # 1. PROCESS PDF
    if st.session_state.master_df is None:
        with st.spinner("Analyzing layout and reconstructing grid..."):
            try:
                # Convert PDF to image (First page only for speed demo)
                images = convert_from_bytes(uploaded_file.read())
                
                # Process the first page (loop this if you need multi-page)
                df = process_image_to_grid(images[0])
                st.session_state.master_df = df
            except Exception as e:
                st.error(f"Error: {e}")

    df = st.session_state.master_df

    if df is not None:
        st.success("Layout Reconstructed!")

        # -------------------------------------------------------
        # TAB 1: PREVIEW & DOWNLOAD (Layout Focus)
        # -------------------------------------------------------
        st.subheader("1. Preview & Download Excel")
        st.caption("This is the raw data matching your visual layout.")
        
        # Display editable grid
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        # DOWNLOAD BUTTON
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, header=False)
        
        st.download_button(
            label="📥 Download Excel File (Original Layout)",
            data=output.getvalue(),
            file_name="scanned_layout.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # -------------------------------------------------------
        # TAB 2: SORT & SUBTOTAL (Data Focus)
        # -------------------------------------------------------
        st.subheader("2. Sort and Calculate")
        st.markdown("Now that you have the columns, you can perform your calculations here.")

        c1, c2, c3 = st.columns(3)
        
        with c1:
            # Sorting
            sort_col = st.selectbox("Sort by Column:", ["None"] + list(edited_df.columns))
        
        with c2:
            # Grouping
            group_col = st.selectbox("Group Subtotals by:", ["None"] + list(edited_df.columns))
            
        with c3:
            # Summing
            sum_col = st.selectbox("Sum Values in:", ["None"] + list(edited_df.columns))

        # LOGIC
        final_view = edited_df.copy()

        # Apply Sort
        if sort_col != "None":
            final_view = final_view.sort_values(by=sort_col)

        # Apply Subtotal
        if group_col != "None" and sum_col != "None":
            try:
                # Clean numbers first
                final_view[sum_col] = (
                    final_view[sum_col].astype(str)
                    .str.replace(r'[^\d\.\-]', '', regex=True)
                )
                final_view[sum_col] = pd.to_numeric(final_view[sum_col], errors='coerce').fillna(0)

                # Calculate GroupBy
                subtotals = final_view.groupby(group_col)[sum_col].sum().reset_index()
                subtotals.columns = [group_col, f"Total {sum_col}"]
                
                st.write("### Subtotal Results")
                st.dataframe(subtotals, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Could not calculate: {e}. Make sure the 'Sum' column has numbers.")

        elif sort_col != "None":
            st.write("### Sorted Data")
            st.dataframe(final_view, use_container_width=True)

    # RESET
    if st.button("Start Over / New File"):
        st.session_state.master_df = None
        st.rerun()
