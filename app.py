import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
import zipfile
import os
import re
import tempfile

# Streamlit app
st.title("Excel to Word Document Generator")
st.write("Upload an Excel file and a Word template to generate Word documents for each product.")

# File uploaders
excel_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
template_file = st.file_uploader("Choose a Word template (.docx)", type=["docx"])

if excel_file and template_file:
    # Read Excel file
    try:
        df = pd.read_excel(excel_file)
        st.write("Uploaded Excel Data Preview:")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # Save template to temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(template_file.read())
        tmp_path = tmp.name

    # Generate Word documents
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            try:
                doc = DocxTemplate(tmp_path)
                # Sanitize column names for Jinja2
                row_dict = {}
                for col in df.columns:
                    # Replace spaces and special characters with underscores
                    safe_key = re.sub(r'[^\w]', '_', col)
                    row_dict[safe_key] = row[col]
                
                # Sanitize Product Name for filename
                product_name = str(row_dict.get("Product_Name", f"document_{index}"))
                safe_filename = re.sub(r'[^\w\s-]', '', product_name).strip().replace(' ', '_')
                
                # Render document
                doc.render(row_dict)
                
                # Save to buffer
                doc_buffer = BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)
                
                # Add to ZIP
                zip_file.writestr(f"{safe_filename}.docx", doc_buffer.getvalue())
            except Exception as e:
                st.warning(f"Error generating document for row {index + 2}: {e}")
                continue

    # Clean up temporary file
    os.unlink(tmp_path)
    
    zip_buffer.seek(0)
    
    # Provide download button
    st.download_button(
        label="Download All Word Documents (ZIP)",
        data=zip_buffer,
        file_name="product_documents.zip",
        mime="application/zip"
    )
    st.success("Word documents generated! Click the button to download the ZIP file.")
elif excel_file or template_file:
    st.warning("Please upload both an Excel file and a Word template to proceed.")