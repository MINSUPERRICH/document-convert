import streamlit as st
import pandas as pd
import pdfplumber
import io

# 1. Page Setup
st.set_page_config(page_title="PDF/Scan to Excel & Subtotal", layout="wide")
st.title("📄 PDF to Excel Converter with Subtotals")

# 2. File Upload
uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

def extract_table_from_pdf(file):
    """
    Attempts to extract the largest table from the first page of the PDF.
    """
    with pdfplumber.open(file) as pdf:
        # Look at the first page (you can loop through pages if needed)
        page = pdf.pages[0] 
        # Extract table
        table = page.extract_table()
        
    if table:
        # Convert list of lists to DataFrame
        # Assume first row is header
        df = pd.DataFrame(table[1:], columns=table[0])
        return df
    else:
        return None

# 3. Main Logic
if uploaded_file is not None:
    st.write("Processing file...")
    
    # Attempt extraction
    try:
        df = extract_table_from_pdf(uploaded_file)
        
        if df is None:
            st.error("Could not detect a table. Is this a scanned image? (See note below)")
        else:
            st.success("Table extracted successfully!")
            
            # 4. Data Cleaning & Display
            st.subheader("1. Review and Edit Data")
            st.markdown("Use the table below to fix any conversion errors before calculating.")
            
            # Convert columns to numeric if possible (for calculations)
            # This logic tries to auto-convert numbers
            for col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='ignore')

            # Allow user to edit data directly in the browser
            edited_df = st.data_editor(df, num_rows="dynamic")

            st.divider()

            # 5. Subtotal Calculation Logic
            st.subheader("2. Calculate Subtotals")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # User picks the "Category" column (e.g., 'Department' or 'Date')
                group_col = st.selectbox("Select column to group by (Category):", edited_df.columns)
            
            with col2:
                # User picks the "Value" column (e.g., 'Amount' or 'Price')
                # We filter to only show numeric columns
                numeric_cols = edited_df.select_dtypes(include=['float64', 'int64']).columns
                value_col = st.selectbox("Select column to sum (Value):", numeric_cols)

            if group_col and value_col:
                # Perform the calculation
                summary_df = edited_df.groupby(group_col)[value_col].sum().reset_index()
                
                # Format for display (add subtotal label)
                summary_df.columns = [group_col, f"Total {value_col}"]
                
                st.write("### Results")
                st.dataframe(summary_df)

                # 6. Download Button
                st.divider()
                st.subheader("3. Download Result")
                
                # Create Excel in memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    edited_df.to_excel(writer, sheet_name='Raw Data', index=False)
                    summary_df.to_excel(writer, sheet_name='Subtotals', index=False)
                
                st.download_button(
                    label="Download Excel File",
                    data=output.getvalue(),
                    file_name="converted_data_with_subtotals.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")

else:
    st.info("Please upload a PDF file containing a table to get started.")
    
    st.markdown("""
    **Note on Scanned Files:** If your PDF is a *picture* of a document (a flat scan), standard extraction won't work. 
    You will need to integrate OCR (Tesseract) which is more complex to set up.
    """)