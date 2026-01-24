import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import numpy as np
import re

# ------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------
st.set_page_config(page_title="AI Data Tool", layout="wide")
st.title("🧠 Smart Scan & Analyze Tool")

# Create two tabs
tab1, tab2 = st.tabs(["📄 Scan PDF to Excel", "🤖 Analyze Excel with AI"])

# ------------------------------------------------------------------
# HELPER: SMART TEXT PARSER (The "Brain")
# ------------------------------------------------------------------
def parse_instruction_and_calculate(df, prompt):
    """
    Reads a text prompt and tries to execute pandas logic.
    Supports: Sum, Average, Count, Sort, Group By.
    """
    prompt = prompt.lower()
    
    # 1. Identify Columns mentioned in the prompt
    # We create a map of {lowercase_name: real_column_name}
    col_map = {c.lower(): c for c in df.columns}
    
    # Find which columns the user typed
    found_cols = [col_map[c] for c in col_map if c in prompt]
    
    # 2. Identify Operation
    op = None
    if any(x in prompt for x in ['sum', 'total', 'add', 'plus']): op = 'sum'
    elif any(x in prompt for x in ['avg', 'average', 'mean']): op = 'mean'
    elif any(x in prompt for x in ['count', 'number of', 'how many']): op = 'count'
    elif 'sort' in prompt: op = 'sort'
    
    # 3. Execution Logic
    try:
        # CASE A: SORTING (e.g., "Sort by Date")
        if op == 'sort':
            if found_cols:
                target_col = found_cols[0] # Pick the first column found
                # Check for ascending/descending
                ascending = False if 'desc' in prompt or 'high to low' in prompt else True
                result = df.sort_values(by=target_col, ascending=ascending)
                return result, f"✅ Sorted data by **{target_col}**"
            else:
                return None, "⚠️ I understood you want to sort, but I couldn't find the column name in your sentence."

        # CASE B: GROUP BY CALCULATION (e.g. "Sum of Price by Color")
        elif op in ['sum', 'mean', 'count']:
            
            # We need to find the Numeric Column (Value) and the Group Column (Category)
            # If the user says "by [Column]", that is likely the grouper.
            
            if 'by' in prompt and len(found_cols) >= 2:
                # Heuristic: split the sentence at 'by'. 
                # Words before 'by' are likely the Value. Words after 'by' are likely the Group.
                parts = prompt.split('by')
                before_by = parts[0]
                after_by = parts[1]
                
                # Find column mentioned before "by"
                val_col = next((c for c in found_cols if c.lower() in before_by), None)
                # Find column mentioned after "by"
                group_col = next((c for c in found_cols if c.lower() in after_by), None)
                
                if val_col and group_col:
                    # Clean the number column first
                    clean_df = df.copy()
                    # Remove '$', ',' and force to number
                    clean_df[val_col] = pd.to_numeric(
                        clean_df[val_col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), 
                        errors='coerce'
                    ).fillna(0)

                    # Perform GroupBy
                    if op == 'sum': 
                        res = clean_df.groupby(group_col)[val_col].sum()
                    elif op == 'mean': 
                        res = clean_df.groupby(group_col)[val_col].mean()
                    elif op == 'count': 
                        res = clean_df.groupby(group_col)[val_col].count()
                    
                    # Format Result
                    res_df = res.reset_index()
                    res_df.columns = [group_col, f"{op.title()} of {val_col}"]
                    return res_df, f"✅ Calculated **{op}** of **{val_col}** grouped by **{group_col}**."
            
            # CASE C: SIMPLE TOTAL (e.g. "Total of Price")
            elif len(found_cols) >= 1:
                target_col = found_cols[0]
                # Clean Data
                clean_series = pd.to_numeric(
                    df[target_col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), 
                    errors='coerce'
                ).fillna(0)
                
                if op == 'sum': val = clean_series.sum()
                elif op == 'mean': val = clean_series.mean()
                elif op == 'count': val = clean_series.count()
                
                # Return as a tiny dataframe for consistency
                return pd.DataFrame({f"{op.title()} of {target_col}": [val]}), f"✅ Calculated total **{op}** of **{target_col}**."

    except Exception as e:
        return None, f"❌ An error occurred while calculating: {e}"

    return None, "❓ I didn't understand. Try phrasing it like: '**Sum** of **Amount** by **Category**' or '**Sort** by **Date**'."

# ------------------------------------------------------------------
# TAB 1: PDF SCANNER (Existing Code)
# ------------------------------------------------------------------
with tab1:
    st.header("1. Scan PDF to Excel")
    
    # --- HELPER: PIXEL ALIGNMENT ALGORITHM ---
    def process_layout_preserving(image, clustering_sensitivity=15):
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
        df = pd.DataFrame(data)
        df = df[df['text'].str.strip() != '']
        if df.empty: return pd.DataFrame()
        
        df['row_id'] = (df['top'] / 15).round().astype(int) 
        all_lefts = df['left'].sort_values().unique()
        col_definitions = [] 
        for x in all_lefts:
            found = False
            for i, c in enumerate(col_definitions):
                if abs(x - c) < clustering_sensitivity:
                    col_definitions[i] = (c + x) / 2
                    found = True; break
            if not found: col_definitions.append(x)
        col_definitions.sort()
        df['col_idx'] = df['left'].apply(lambda x: np.argmin([abs(x - c) for c in col_definitions]))
        
        grid = [['' for _ in range(len(col_definitions))] for _ in range(len(df['row_id'].unique()))]
        row_map = {rid: i for i, rid in enumerate(sorted(df['row_id'].unique()))}
        for _, row in df.iterrows():
            grid[row_map[row['row_id']]][row['col_idx']] = (grid[row_map[row['row_id']]][row['col_idx']] + " " + row['text']).strip()
            
        final = pd.DataFrame(grid)
        final = final.loc[:, (final != '').any(axis=0)]
        final.columns = [f"Col_{i+1}" for i in range(final.shape[1])]
        return final

    if 'scan_df' not in st.session_state: st.session_state.scan_df = None
    uploaded_pdf = st.file_uploader("Upload Scanned PDF", type=["pdf"], key="pdf1")
    
    if uploaded_pdf:
        with st.expander("Alignment Settings"):
            sens = st.slider("Sensitivity", 5, 100, 25)
        try:
            images = convert_from_bytes(uploaded_pdf.read())
            st.session_state.scan_df = process_layout_preserving(images[0], sens)
            st.success("Scan Processed!")
            edited_scan = st.data_editor(st.session_state.scan_df, num_rows="dynamic", use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer: edited_scan.to_excel(writer, index=False, header=False)
            st.download_button("📥 Download Excel", output.getvalue(), "scan.xlsx")
        except Exception as e: st.error(f"Error: {e}")

# ------------------------------------------------------------------
# TAB 2: AI ANALYZER (New Logic)
# ------------------------------------------------------------------
with tab2:
    st.header("2. Ask the App to Calculate")
    st.caption("Upload an Excel file and type instructions like: 'Sum of Amount by Category'")
    
    uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], key="xls1")
    
    if uploaded_excel:
        df_excel = pd.read_excel(uploaded_excel)
        
        # 1. SHOW DATA
        with st.expander("👀 View Raw Data", expanded=True):
            st.dataframe(df_excel.head())
            st.markdown(f"**Available Columns:** {', '.join(df_excel.columns)}")

        st.divider()

        # 2. THE CHAT INTERFACE
        st.subheader("What do you want to calculate?")
        
        col_input, col_btn = st.columns([4, 1])
        with col_input:
            user_query = st.text_input("Type your direction here:", placeholder="e.g., Calculate total Price by Product")
        with col_btn:
            st.write("") # Spacer
            st.write("") # Spacer
            run_calc = st.button("🚀 Run", type="primary")

        # 3. PROCESS THE REQUEST
        if run_calc and user_query:
            result_df, message = parse_instruction_and_calculate(df_excel, user_query)
            
            st.info(message) # Show what the app understood
            
            if result_df is not None:
                st.dataframe(result_df, use_container_width=True)
                
                # Download Result
                out_buff = io.BytesIO()
                with pd.ExcelWriter(out_buff, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Download Result",
                    data=out_buff.getvalue(),
                    file_name="calculation_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
