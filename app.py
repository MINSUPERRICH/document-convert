import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import numpy as np
import re
from decimal import Decimal, ROUND_HALF_UP

# ------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------
st.set_page_config(page_title="AI Data Tool", layout="wide")
st.title("🧠 Smart Scan & Analyze Tool")

tab1, tab2 = st.tabs(["📄 Scan PDF to Excel", "🤖 Analyze Excel with AI"])

# ------------------------------------------------------------------
# HELPER: PRECISION ROUNDING (The Fix)
# ------------------------------------------------------------------
def strict_invoice_round(x):
    """
    Rounds a number to exactly 2 decimal places using standard arithmetic rounding.
    Fixes 'Double Rounding' bugs by converting the raw value directly.
    """
    try:
        # Check for empty or non-numeric
        if pd.isna(x) or str(x).strip() == "":
            return 0.0
        
        # Convert directly to string to preserve raw precision (e.g. "152.734999")
        # Then convert to Decimal
        d = Decimal(str(x))
        
        # Round to 0.01 using ROUND_HALF_UP (Standard School/Excel Method)
        # This ensures 152.734999 -> 152.73 (Correct)
        # Whereas previously it might have bumped to 152.74 (Incorrect)
        val = float(d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
        return val
    except:
        return 0.0

# ------------------------------------------------------------------
# TAB 1: SCAN PDF (OCR Logic)
# ------------------------------------------------------------------
with tab1:
    st.header("1. Scan PDF to Excel")
    
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
# TAB 2: AI ANALYZER (Precision Mode)
# ------------------------------------------------------------------
with tab2:
    st.header("2. Ask the App to Calculate")
    
    uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], key="xls1")
    
    if uploaded_excel:
        st.info("Check the preview. Adjust 'Header Row Number' if needed.")
        header_row_idx = st.number_input("Header Row Number", min_value=0, max_value=20, value=0)
        
        # Load Data
        df_excel = pd.read_excel(uploaded_excel, header=header_row_idx)

        # 1. APPLY PRECISION ROUNDING IMMEDIATELY
        # This ensures the data in memory exactly matches what you see on paper
        numeric_cols = df_excel.select_dtypes(include=['float', 'float64']).columns
        for col in numeric_cols:
             df_excel[col] = df_excel[col].apply(strict_invoice_round)

        with st.expander("👀 View Data Preview", expanded=True):
            st.dataframe(df_excel.head())
            st.caption(f"**Valid Column Names:** {', '.join(list(df_excel.columns))}")

        st.divider()

        st.subheader("What do you want to calculate?")
        user_query = st.text_input("Direction:", placeholder="e.g. Sort by Date. Sum all columns by Class.")
        run_calc = st.button("🚀 Run", type="primary")

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
                    elif any(x in step_lower for x in ['sum', 'total', 'add', 'subtotal']): op = 'sum'
                    elif any(x in step_lower for x in ['avg', 'average']): op = 'mean'
                    elif any(x in step_lower for x in ['count']): op = 'count'
                    
                    # EXECUTE
                    if op == 'sort' and found_cols:
                        target = found_cols[0]
                        ascending = False if 'desc' in step_lower else True
                        current_df = current_df.sort_values(by=target, ascending=ascending)
                        log_messages.append(f"✅ Sorted by **{target}**")
                        final_result = current_df 

                    elif op in ['sum', 'mean', 'count']:
                        
                        group_col = None
                        if 'by' in step_lower:
                            parts = step_lower.split('by')
                            group_candidates = [c for c in found_cols if c.lower() in parts[1]]
                            if group_candidates: group_col = group_candidates[0]
                        
                        if group_col:
                            # Identify Values
                            val_cols = [c for c in found_cols if c != group_col]
                            if 'all' in step_lower or 'each' in step_lower or not val_cols:
                                target_cols = current_df.select_dtypes(include=[np.number]).columns.tolist()
                                if group_col in target_cols: target_cols.remove(group_col)
                            else:
                                target_cols = val_cols

                            if target_cols:
                                # Ensure we are calculating on strictly rounded numbers
                                for c in target_cols:
                                     # Clean string artifacts if any
                                     current_df[c] = pd.to_numeric(
                                        current_df[c].astype(str).str.replace(r'[^\d\.\-]', '', regex=True),
                                        errors='coerce'
                                     )
                                     # Force Rounding One Last Time Before Sum
                                     current_df[c] = current_df[c].apply(strict_invoice_round)

                                res = current_df.groupby(group_col)[target_cols].agg(op)
                                res_df = res.reset_index()
                                res_df.columns = [group_col] + [f"{op.title()} of {c}" for c in target_cols]
                                final_result = res_df
                                log_messages.append(f"✅ Calculated **{op}** for **{len(target_cols)} columns** by **{group_col}**")
                            else:
                                log_messages.append("⚠️ Found group column but no values to sum.")
                        else:
                            # Simple Total
                            if found_cols:
                                target = found_cols[0]
                                current_df[target] = pd.to_numeric(current_df[target].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), errors='coerce')
                                current_df[target] = current_df[target].apply(strict_invoice_round)
                                
                                val = current_df[target].agg(op)
                                final_result = pd.DataFrame({f"{op.title()} of {target}": [val]})
                                log_messages.append(f"✅ Calculated total **{op}** of **{target}**")

            except Exception as e:
                st.error(f"Something went wrong: {e}")

            # FINAL DISPLAY
            for msg in log_messages:
                st.write(msg)
            
            if final_result is not None:
                # Apply rounding to the final result (just to be safe)
                res_numeric = final_result.select_dtypes(include=['float', 'float64']).columns
                for c in res_numeric:
                    final_result[c] = final_result[c].apply(strict_invoice_round)

                st.dataframe(final_result, use_container_width=True)
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as writer:
                    final_result.to_excel(writer, index=False)
                st.download_button("📥 Download Result", out.getvalue(), "result.xlsx")
            else:
                st.warning("I processed the steps but couldn't generate a result.")
