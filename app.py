import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import numpy as np

# ------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------
st.set_page_config(page_title="Pixel-Perfect Invoice Scanner", layout="wide")
st.title("📄 Pixel-Perfect Invoice to Excel")

# ------------------------------------------------------------------
# ALGORITHM: VISUAL COLUMN CLUSTERING
# ------------------------------------------------------------------

def process_layout_preserving(image, clustering_sensitivity=15):
    """
    1. Detects words and their X (horizontal) positions.
    2. Clusters similar X-positions to define 'Global Columns'.
    3. Maps every row's words into these Global Columns.
    """
    # 1. Get detailed OCR data
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    df = pd.DataFrame(data)
    
    # Filter noise
    df = df[df['text'].str.strip() != '']
    df['text'] = df['text'].astype(str)
    
    if df.empty:
        return pd.DataFrame()

    # 2. Define Rows (Y-Axis)
    # Round Y-coordinates to group words on the same line
    df['row_id'] = (df['top'] / 15).round().astype(int) 

    # 3. Define Global Columns (The Core Logic)
    # We look at the 'left' coordinate of EVERY word in the document.
    # Words that start at similar X positions (e.g., 500px, 502px, 498px) belong to the same column.
    
    # We use a simple clustering approach:
    all_lefts = df['left'].sort_values().unique()
    
    col_definitions = [] # List of representative X-values
    
    for x in all_lefts:
        # Check if this x belongs to an existing cluster
        found_cluster = False
        for i, center in enumerate(col_definitions):
            if abs(x - center) < clustering_sensitivity: # Sensitivity threshold (pixels)
                # Update cluster center (weighted average could be better, but simple average works)
                col_definitions[i] = (center + x) / 2
                found_cluster = True
                break
        
        if not found_cluster:
            col_definitions.append(x)
            
    col_definitions.sort()
    
    # Map every word to a column index (0, 1, 2...)
    def get_col_index(x_val):
        # Find closest column center
        distances = [abs(x_val - c) for c in col_definitions]
        return np.argmin(distances)

    df['col_idx'] = df['left'].apply(get_col_index)

    # 4. Build the Grid
    # Create empty grid: Rows x Columns
    unique_rows = sorted(df['row_id'].unique())
    num_cols = len(col_definitions)
    
    grid = [['' for _ in range(num_cols)] for _ in range(len(unique_rows))]
    
    # Map row_ids to list indices 0..N
    row_map = {rid: i for i, rid in enumerate(unique_rows)}

    for _, row in df.iterrows():
        r_idx = row_map[row['row_id']]
        c_idx = row['col_idx']
        txt = row['text']
        
        # Append text if cell already has content (merging split words)
        if grid[r_idx][c_idx]:
            grid[r_idx][c_idx] += " " + txt
        else:
            grid[r_idx][c_idx] = txt

    # Convert to DataFrame
    final_df = pd.DataFrame(grid)
    
    # Clean up: Drop columns that are completely empty
    final_df = final_df.loc[:, (final_df != '').any(axis=0)]
    
    # Name columns generically
    final_df.columns = [f"Col_{i+1}" for i in range(final_df.shape[1])]
    
    return final_df

# ------------------------------------------------------------------
# MAIN APP
# ------------------------------------------------------------------

if 'scan_df' not in st.session_state:
    st.session_state.scan_df = None

uploaded_file = st.file_uploader("Upload Scanned PDF", type=["pdf"])

if uploaded_file:
    # PREVIEW SETTINGS
    with st.expander("⚙️ Alignment Settings (Open if columns are messy)", expanded=False):
        sensitivity = st.slider(
            "Column Sensitivity (Lower = More Columns, Higher = Merged Columns)", 
            min_value=5, 
            max_value=100, 
            value=25,
            help="If columns are splitting too much (e.g. '$' separate from '100'), increase this number."
        )
        
    # PROCESS
    # We re-run this only if file changes or button pressed, 
    # but for slider responsiveness we run it on change.
    try:
        images = convert_from_bytes(uploaded_file.read())
        # Processing First Page Only for Speed
        df = process_layout_preserving(images[0], clustering_sensitivity=sensitivity)
        st.session_state.scan_df = df
    except Exception as e:
        st.error(f"Processing Error: {e}")

    df = st.session_state.scan_df

    if df is not None:
        st.success("Structure Reconstructed!")

        # -------------------------------------------------------
        # SECTION 1: CLEANUP & DOWNLOAD
        # -------------------------------------------------------
        st.subheader("1. Verify & Download")
        st.caption("Check the 'TOTAL' row at the bottom. It should now align with the Price column.")
        
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        # DOWNLOAD
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, header=False)
        
        col_dl, col_dummy = st.columns([1, 3])
        with col_dl:
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name="aligned_scan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.divider()

        # -------------------------------------------------------
        # SECTION 2: SORT & SUBTOTAL
        # -------------------------------------------------------
        st.subheader("2. Sort & Subtotal")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            sort_col = st.selectbox("Sort by:", ["None"] + list(edited_df.columns))
        with c2:
            group_col = st.selectbox("Group by:", ["None"] + list(edited_df.columns))
        with c3:
            sum_col = st.selectbox("Sum values in:", ["None"] + list(edited_df.columns))

        # LOGIC
        calc_df = edited_df.copy()

        # 1. Clean Data (Remove 'Total' row so it doesn't mess up sorting)
        # We assume any row containing "TOTAL" in the first few columns is a footer
        mask = calc_df.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False, na=False)).any(axis=1)
        data_rows = calc_df[~mask]  # Rows WITHOUT 'Total'
        footer_rows = calc_df[mask] # Rows WITH 'Total'

        # 2. Sorting
        if sort_col != "None":
            data_rows = data_rows.sort_values(by=sort_col)
            st.info(f"Data sorted by {sort_col}")

        # 3. Calculation
        if group_col != "None" and sum_col != "None":
            try:
                # Clean numbers: remove '$', ',', spaces
                clean_vals = data_rows[sum_col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
                data_rows['numeric_val'] = pd.to_numeric(clean_vals, errors='coerce').fillna(0)
                
                # Group
                summary = data_rows.groupby(group_col)['numeric_val'].sum().reset_index()
                summary.columns = [group_col, f"Total {sum_col}"]
                
                st.write("### Subtotal Results")
                st.dataframe(summary, use_container_width=True)
                
            except Exception as e:
                st.error(f"Calculation failed: {e}")
        else:
            # If no calc, just show sorted list
            st.write("### Working Data")
            st.dataframe(data_rows, use_container_width=True)

    if st.button("Reset"):
        st.session_state.scan_df = None
        st.rerun()
