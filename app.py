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

tab1, tab2 = st.tabs(["📄 Scan PDF to Excel", "🤖 Analyze Excel with AI"])

# ------------------------------------------------------------------
# TAB 1: SCAN PDF (OCR Logic)
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
        
        if not col_definitions: return pd.DataFrame()

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
# TAB 2: AI ANALYZER (With Rounding Fix)
# ------------------------------------------------------------------
with tab2:
    st.header("2. Ask the App to Calculate")
    
    uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], key="xls1")
    
    if uploaded_excel:
        # 1. HEADER SELECTION
        st.info("Check the preview below. If headers are wrong, change the 'Header Row Number'.")
        header_row_idx = st.number_input("Header Row Number (0 = First Row)", min_value=0, max_value=20, value=0)
        
        df_excel = pd.read_excel(uploaded_excel, header=header_row_idx)
        
        # Display Preview
        with st.expander("👀 View Data Preview", expanded=True):
            st.dataframe(df_excel.head())
            st.caption(f"Active Columns: {', '.join(list(df_excel.columns))}")

        st.divider()

        # 2. CHAT INTERFACE
        st.subheader("What do you want to calculate?")
        user_query = st.text_input("Direction:", placeholder="e.g. Sort by class, then sum of Amount")
        run_calc = st.button("🚀 Run", type="primary")

        # 3. MULTI-STEP PARSER
        if run_calc and user_query:
            
            steps = re.split(r'[.;]| then | after that | and ', user_query, flags=re.IGNORECASE)
            steps = [s.strip() for s in steps if s.strip()]
            
            current_df = df_excel.copy()
            final_result = None
            log_messages = []
            
            try:
                for step in steps:
                    step_lower = step.lower()
                    
                    # Match Columns
                    col_map = {c.lower(): c for c in current_df.columns}
                    sorted_cols = sorted(col_map.keys(), key=len, reverse=True)
                    found_cols = [col_map[c] for c in sorted_cols if c in step_lower]
                    
                    # Identify Op
                    op = None
                    if 'sort' in step_lower: op = 'sort'
                    elif any(x in step_lower for x in ['sum', 'total', 'add']): op = 'sum'
                    elif any(x in step_lower for x in ['avg', 'average']): op = 'mean'
                    elif any(x in step_lower for x in ['count']): op = 'count'
                    
                    # Execution
                    if op == 'sort' and found_cols:
                        target = found_cols[0]
                        ascending = False if 'desc' in step_lower else True
                        current_df = current_df.sort_values(by=target, ascending=ascending)
                        log_messages.append(f"✅ Sorted by **{target}**")
                        final_result = current_df 

                    elif op in ['sum', 'mean', 'count'] and found_cols:
                        val_col = found_cols[0] 
                        
                        # Clean numbers
                        clean_df = current_df.copy()
                        # Force numeric conversion
                        clean_df[val_col] = pd.to_numeric(
                            clean_df[val_col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True),
                            errors='coerce'
                        ).fillna(0)

                        if 'by' in step_lower and len(found_cols) > 1:
                            # Group By
                            parts = step_lower.split('by')
                            group_candidates = [c for c in found_cols if c.lower() in parts[1]]
                            
                            if group_candidates:
                                group_col = group_candidates[0]
                                val_col = [c for c in found_cols if c != group_col][0]
                                
                                res = clean_df.groupby(group_col)[val_col].agg(op)
                                res_df = res.reset_index()
                                res_df.columns = [group_col, f"{op.title()} of {val_col}"]
                                final_result = res_df
                                log_messages.append(f"✅ Calculated **{op}** of **{val_col}** by **{group_col}**")
                        else:
                            # Simple Total
                            val = clean_df[val_col].agg(op)
                            final_result = pd.DataFrame({f"{op.title()} of {val_col}": [val]})
                            log_messages.append(f"✅ Calculated total **{op}** of **{val_col}**")
                            
                # 4. ROUNDING FIX AND DISPLAY
                for msg in log_messages:
                    st.write(msg)
                
                if final_result is not None:
                    # ---> THE FIX: Round all numeric columns to 2 decimals <---
                    numeric_cols = final_result.select_dtypes(include=['float', 'float64']).columns
                    final_result[numeric_cols] = final_result[numeric_cols].round(2)

                    st.dataframe(final_result, use_container_width=True)
                    
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as writer:
                        final_result.to_excel(writer, index=False)
                    st.download_button("📥 Download Result", out.getvalue(), "result.xlsx")
                else:
                    st.warning("I processed the steps but couldn't generate a result. Please check column names.")

            except Exception as e:
                st.error(f"Something went wrong: {e}")
