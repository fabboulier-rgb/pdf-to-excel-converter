import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="Universal PDF to Excel Converter", page_icon="📊")

st.title("Universal PDF to Excel Converter")
st.markdown("Upload any tabular PDF. The engine will detect distinct tables and export them into a multi-tabbed Excel file.")

# File uploader
uploaded_file = st.file_uploader("Upload PDF Document", type="pdf")

if uploaded_file is not None:
    st.info("File uploaded securely. Click below to begin extraction.")
    
    if st.button("Convert to Excel"):
        with st.spinner("Analyzing document and mapping tables..."):
            extracted_tables = []
            
            try:
                # Open PDF directly from memory
                with pdfplumber.open(uploaded_file) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        # Extract all distinct tables on the current page
                        tables = page.extract_tables()
                        
                        for table_index, table in enumerate(tables):
                            cleaned_table = []
                            for row in table:
                                # Clean line breaks and normalize empty cells
                                clean_row = [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in row]
                                
                                # Skip entirely empty rows
                                if any(clean_row):
                                    cleaned_table.append(clean_row)
                            
                            # If valid data exists, save it with a dynamic page tracker
                            if cleaned_table:
                                # Standardize column counts to prevent Excel corruption
                                max_cols = max(len(r) for r in cleaned_table)
                                normalized_table = [r + [""] * (max_cols - len(r)) for r in cleaned_table]
                                
                                df = pd.DataFrame(normalized_table)
                                tab_name = f"Pg{page_num + 1}_Tbl{table_index + 1}"
                                extracted_tables.append((tab_name, df))
                                
                if extracted_tables:
                    # Write all tables to an in-memory Excel file
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        for tab_name, df in extracted_tables:
                            # Write without headers/index to preserve raw PDF layout
                            df.to_excel(writer, sheet_name=tab_name[:31], index=False, header=False)
                    
                    st.success(f"Success! Extracted {len(extracted_tables)} distinct tables.")
                    
                    # Provide secure download link
                    st.download_button(
                        label="⬇️ Download Excel File",
                        data=excel_buffer.getvalue(),
                        file_name="Converted_Tables.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No structured grid tables could be detected in this PDF. It may be a scanned image.")
                    
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
