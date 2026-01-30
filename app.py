import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import numpy as np
import re
from decimal import Decimal, ROUND_HALF_UP
import google.generativeai as genai
from docx import Document

# ------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------
st.set_page_config(page_title="AI Data Tool", layout="wide")
st.title("🧠 Smart Scan & Analyze Tool")

# Create 3 Tabs
tab1, tab2, tab3 = st.tabs(["📄 Scan PDF to Excel", "🤖 Analyze Excel", "📝 Generate Reports (Docs)"])

# ------------------------------------------------------------------
# SHARED HELPER FUNCTIONS
# ------------------------------------------------------------------
def strict_invoice_round(x):
    """Rounds to exactly 2 decimal places (Standard Rounding)"""
    try:
        if pd.isna(x) or str(x).strip() == "": return 0.0
        d = Decimal(str(x))
        return float(d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except: return 0.0

def process_layout_preserving(image, clustering_sensitivity=15):
    """Core OCR Logic for Tables"""
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

# ------------------------------------------------------------------
# TAB 1: SCAN PDF TO EXCEL
# ------------------------------------------------------------------
with tab1:
    st.header("1. Scan PDF to Excel")
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
# TAB 2: ANALYZE EXCEL
# ------------------------------------------------------------------
with tab2:
    st.header("2. Ask the App to Calculate")
    uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], key="xls1")
    
    if uploaded_excel:
        st.info("Check preview. Adjust Header Row if needed.")
        header_row_idx = st.number_input("Header Row Number", min_value=0, max_value=20, value=0)
        df_excel = pd.read_excel(uploaded_excel, header=header_row_idx)
        
        with st.expander("👀 View Data Preview", expanded=True):
            st.dataframe(df_excel.head().round(2))
            st.caption(f"Valid Columns: {', '.join(list(df_excel.columns))}")
        
        st.divider()
        user_query = st.text_input("Direction:", placeholder="e.g. Sort by Date. Sum all columns by Class.")
        run_calc = st.button("🚀 Run Analysis", type="primary")

        if run_calc and user_query:
            steps = re.split(r'[.;]| then | after that | and ', user_query, flags=re.IGNORECASE)
            steps = [s.strip() for s in steps if s.strip()]
            current_df = df_excel.copy()
            final_result = None
            log_messages = []
            
            try:
                for step in steps:
                    step_lower = step.lower()
                    col_map = {c.lower(): c for c in current_df.columns}
                    sorted_cols = sorted(col_map.keys(), key=len, reverse=True)
                    found_cols = [col_map[c] for c in sorted_cols if c in step_lower]
                    
                    op = None
                    if 'sort' in step_lower: op = 'sort'
                    elif any(x in step_lower for x in ['sum', 'total', 'add', 'subtotal']): op = 'sum'
                    elif any(x in step_lower for x in ['avg', 'average']): op = 'mean'
                    elif any(x in step_lower for x in ['count']): op = 'count'
                    
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
                            cands = [c for c in found_cols if c.lower() in parts[1]]
                            if cands: group_col = cands[0]
                        
                        if group_col:
                            val_cols = [c for c in found_cols if c != group_col]
                            if 'all' in step_lower or 'each' in step_lower or not val_cols:
                                target_cols = current_df.select_dtypes(include=[np.number]).columns.tolist()
                                if group_col in target_cols: target_cols.remove(group_col)
                            else: target_cols = val_cols

                            if target_cols:
                                for c in target_cols:
                                     current_df[c] = pd.to_numeric(current_df[c].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), errors='coerce')
                                res = current_df.groupby(group_col)[target_cols].agg(op)
                                res_df = res.reset_index()
                                res_df.columns = [group_col] + [f"{op.title()} of {c}" for c in target_cols]
                                final_result = res_df
                                log_messages.append(f"✅ Calculated **{op}** for **{len(target_cols)} columns** by **{group_col}**")
                        else:
                            if found_cols:
                                target = found_cols[0]
                                current_df[target] = pd.to_numeric(current_df[target].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), errors='coerce')
                                val = current_df[target].agg(op)
                                final_result = pd.DataFrame({f"{op.title()} of {target}": [val]})
                                log_messages.append(f"✅ Calculated total **{op}** of **{target}**")
            except Exception as e: st.error(f"Error: {e}")

            for msg in log_messages: st.write(msg)
            if final_result is not None:
                res_numeric = final_result.select_dtypes(include=['float', 'float64']).columns
                disp = final_result.copy()
                disp[res_numeric] = disp[res_numeric].round(2)
                st.dataframe(disp, use_container_width=True)
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as writer: disp.to_excel(writer, index=False)
                st.download_button("📥 Download Result", out.getvalue(), "result.xlsx")

# ------------------------------------------------------------------
# TAB 3: GENERATE REPORTS (DOCS)
# ------------------------------------------------------------------
with tab3:
    st.header("3. Generate Word/Google Docs")
    st.markdown("Use this tab to create **Formal Reports** or **Clean Tables** from your PDF using AI.")
    
    # API KEY INPUT
    with st.expander("🔑 Setup AI (Required for Summary)", expanded=True):
        api_key = st.text_input("Enter Google Gemini API Key:", type="password", help="Get a free key from Google AI Studio")
        st.caption("Don't have a key? [Get one here](https://aistudio.google.com/app/apikey) (It's free).")

    uploaded_doc_pdf = st.file_uploader("Upload PDF for Report", type=["pdf"], key="pdf_docs")
    
    if uploaded_doc_pdf:
        # 1. READ TEXT (Basic OCR)
        # We use a simpler OCR extraction here because we want continuous text for the AI, not just a grid
        images = convert_from_bytes(uploaded_doc_pdf.read())
        raw_text = ""
        for img in images:
            raw_text += pytesseract.image_to_string(img) + "\n"
            
        st.success("PDF Loaded successfully!")
        
        # 2. CHOOSE TEMPLATE
        st.subheader("Choose a Document Style")
        template_choice = st.radio(
            "What do you want to create?", 
            ["Option 1: Data Extraction (Best for Excel)", "Option 2: Research Summary (Best for Word)"]
        )
        
        # PREPARE PROMPTS
        if "Option 1" in template_choice:
            default_prompt = (
                "I am uploading a PDF. Please extract the financial data/lists and present them in a Markdown table. "
                "Ensure every row is filled and do not include any conversational text before or after the table. "
                "Give me the table only."
            )
        else:
            default_prompt = (
                "Please summarize this PDF into a formal report. Use Markdown headers (##) for each section, "
                "use bullet points for key takeaways, and include a 'Conclusion' at the end. "
                "Format this so it is ready to be exported into a professional document."
            )

        user_prompt = st.text_area("Customize Instructions (Optional):", value=default_prompt, height=100)
        
        # 3. GENERATE
        if st.button("✨ Generate Document"):
            if not api_key:
                st.warning("⚠️ Please enter a Google Gemini API Key at the top to use this feature.")
            else:
                try:
                    with st.spinner("AI is reading and writing your report..."):
                        # Configure Gemini
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-1.5-flash')
                        
                        # Send Request
                        full_prompt = f"{user_prompt}\n\nHere is the document text:\n{raw_text}"
                        response = model.generate_content(full_prompt)
                        generated_text = response.text
                        
                        st.subheader("Generated Result")
                        st.markdown(generated_text)
                        
                        # 4. EXPORT TO WORD
                        doc = Document()
                        doc.add_heading('Generated Report', 0)
                        
                        # Simple markdown-to-word parser
                        for line in generated_text.split('\n'):
                            clean_line = line.strip()
                            if clean_line.startswith('## '):
                                doc.add_heading(clean_line.replace('## ', ''), level=2)
                            elif clean_line.startswith('# '):
                                doc.add_heading(clean_line.replace('# ', ''), level=1)
                            elif clean_line.startswith('* ') or clean_line.startswith('- '):
                                doc.add_paragraph(clean_line.replace('* ', '').replace('- ', ''), style='List Bullet')
                            else:
                                if clean_line: doc.add_paragraph(clean_line)

                        # Save to buffer
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        st.download_button(
                            label="📥 Download as Word Doc (.docx)",
                            data=doc_io,
                            file_name="generated_report.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                except Exception as e:
                    st.error(f"AI Error: {e}. Check your API Key.")
