import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import re
from PIL import Image, ImageOps

# ------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------
st.set_page_config(page_title="Smart Invoice to Excel", layout="wide")
st.title("📄 Smart Invoice/Scan to Excel")

# ------------------------------------------------------------------
# HELPER: IMAGE PROCESSING
# ------------------------------------------------------------------

def process_image_for_ocr(image):
    """
    Makes the image black & white and 2x larger.
    This helps OCR read small invoice numbers drastically better.
    """
    # 1. Grayscale
    image = ImageOps.grayscale(image)
    # 2. Resize (Double resolution)
    width, height = image.size
    image = image.resize((width * 2, height * 2), Image.Resampling.LANCZOS)
    return image

def extract_text_from_pdf(file_bytes):
    """ Converts PDF to Image -> Pre-process -> OCR """
    try:
        images = convert_from_bytes(file_bytes)
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return ""

    full_text = ""
    progress = st.progress(0)
    
    for i, img in enumerate(images):
        # Optimize image for tables
        processed_img = process_image_for_ocr(img)
        
        # --psm 6: Assume a single uniform block of text
        text = pytesseract.image_to_string(processed_img, config='--psm 6')
        full_text += text + "\n"
        progress.progress((i + 1) / len(images))
        
    progress.empty()
    return full_text

# ------------------------------------------------------------------
# HELPER: PARSING LOGIC (THE FIX)
# ------------------------------------------------------------------

def parse_invoice_mode(text):
    """
    SMART MODE: Instead of splitting by space, we look for the Price 
    at the end of the line and assume everything else is description.
    """
    lines = text.split('\n')
    data = []

    # Regex to find a price at the end of a line 
    # Examples: "1,234.56" or "500.00" or "$123.45"
    # It looks for: Numbers + optional comma + dot + 2 digits + End of String
    price_pattern = re.compile(r'[\$]?([0-9,]+\.[0-9]{2})$')

    for line in lines:
        clean_line = line.strip()
        if not clean_line:
            continue
            
        # 1. Search for price at the end
        match = price_pattern.search(clean_line)
        
        if match:
            # We found a Price! 
            price_str = match.group(1) # The number part
            
            # The description is everything BEFORE the price
            # We remove the price from the line to get the rest
            description_part = clean_line[:match.start()].strip()
            
            # Clean up the price (remove commas for math)
            numeric_value = float(price_str.replace(',', ''))
            
            data.append({
                "Full_Row_Text": description_part,  # Keep context
                "Extracted_Value": numeric_value
            })
        else:
            # If no price found, we add it as a "Text Only" row (header or notes)
            # We skip short noise (like "4W" on its own if it has no price)
            if len(clean_line) > 3:
                data.append({
                    "Full_Row_Text": clean_line,
                    "Extracted_Value": 0.0
                })

    return pd.DataFrame(data)

def parse_simple_mode(text):
    """ ORIGINAL MODE: Split by spaces (Good for simple lists) """
    lines = text.split('\n')
    data = []
    for line in lines:
        if line.strip():
            row = re.split(r'\s{2,}', line.strip()) # Split by 2+ spaces
            data.append(row)
    
    # Normalize
    if data:
        max_cols = max(len(r) for r in data)
        data = [r + [''] * (max_cols - len(r)) for r in data]
        return pd.DataFrame(data, columns=[f"Col_{i}" for i in range(max_cols)])
    return pd.DataFrame()

# ------------------------------------------------------------------
# MAIN APP
# ------------------------------------------------------------------

if 'ocr_cache' not in st.session_state:
    st.session_state.ocr_cache = ""

uploaded_file = st.file_uploader("Upload PDF Invoice/Scan", type=["pdf"])

if uploaded_file:
    # 1. RUN OCR (If new file)
    if st.session_state.ocr_cache == "":
        st.info("Scanning & Enhancing Image...")
        file_bytes = uploaded_file.read()
        st.session_state.ocr_cache = extract_text_from_pdf(file_bytes)

    raw_text = st.session_state.ocr_cache

    if raw_text:
        st.divider()
        
        # 2. CHOOSE MODE
        st.subheader("1. Extraction Strategy")
        mode = st.radio(
            "Select how to read this file:",
            ["Smart Invoice Mode (Best for your file)", "Simple Table Mode"],
            horizontal=True
        )
        
        st.caption("ℹ️ **Smart Mode** looks for a price at the end of every line (e.g. 3,780.89) and aligns the row automatically.")

        # 3. PARSE DATA
        if mode.startswith("Smart"):
            df = parse_invoice_mode(raw_text)
            
            # Show the result immediately
            st.subheader("2. Verify Data")
            
            # Allow user to edit
            edited_df = st.data_editor(
                df, 
                num_rows="dynamic", 
                use_container_width=True,
                column_config={
                    "Extracted_Value": st.column_config.NumberColumn(
                        "Amount",
                        format="$%.2f"
                    ),
                    "Full_Row_Text": st.column_config.TextColumn(
                        "Description / Details",
                        width="large"
                    )
                }
            )
            
            # CALCULATE
            st.divider()
            st.subheader("3. Subtotals")
            
            # Since we already separated Description and Value, we just need to Group
            # However, for Invoices, usually you just want the Grand Total or filter by Text
            
            total_sum = edited_df["Extracted_Value"].sum()
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric(label="Grand Total", value=f"${total_sum:,.2f}")
            
            # Export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, sheet_name='Invoice Data', index=False)
            
            st.download_button("📥 Download Excel", output.getvalue(), "invoice_smart_export.xlsx")

        else:
            # Fallback to the old way if Smart Mode fails
            df = parse_simple_mode(raw_text)
            st.write("Raw Columns Detected:")
            st.data_editor(df)
            st.warning("Switch to 'Smart Invoice Mode' above for better results with prices.")

    if st.button("Reset / New File"):
        st.session_state.ocr_cache = ""
        st.rerun()
