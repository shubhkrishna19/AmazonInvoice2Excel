#!/usr/bin/env python3

"""
Amazon Invoices PDF to Single Excel Converter - Streamlit Version
Extracts invoice data from Amazon PDF invoices into organized Excel format
"""

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
from io import BytesIO

class FinalPerfectExtractor:
    def __init__(self):
        self.extraction_patterns = {
            'Order Number': [
                r'Order Number:\s*([\d-]+)',
                r'Order.*?([\d]{3}-[\d]{10}-[\d]{7})',
            ],
            'Order Date': [
                r'Order Date:\s*([\d./]+)',
                r'([\d]{2}\.[\d]{2}\.[\d]{4})',
            ],
            'Invoice Number': [
                r'Invoice Number\s*:\s*([A-Z]+-[\d]+)',
                r'([A-Z]{3,4}-[\d]+)',
                r'\b(BLX1-\d+)\b',
                r'\b(IN-\d+)\b',
            ],
            'Invoice Details': [
                r'Invoice Details\s*:\s*([A-Z]+-[A-Z]+-[\d]+-[\d]+)',
                r'(UP-[A-Z]+-[\d-]+)',
                r'(KA-[A-Z]+-[\d-]+)',
                r'\bUP-143350511-\d+\b',
                r'\bKA-BLX1-[\d-]+\b',
            ],
            'Customer Address': [
                r'Shipping Address[\s:]*([\s\S]+?)(?=\nState/UT Code)',
                r'Shipping Address[\s:]*([\s\S]+?)(?=State/UT|Place of)',
            ],
            'Total Amount': [
                r'Invoice Value[:\s]*₹?\s*([\d,]+\.00)',  # Only .00 amounts
            ],
        }

    def extract_text_from_pdf_bytes(self, pdf_bytes):
        """Extract text from PDF bytes with error handling"""
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            full_text = ""
            for page in doc:
                text = page.get_text()
                full_text += text + "\n"
            doc.close()
            return full_text
        except Exception as e:
            st.error(f"Error extracting text from PDF: {e}")
            return ""

    def extract_description_from_table(self, text):
        """Extract description from the largest table cell containing BLUEWUD product info"""
        description_patterns = [
            r'\d+\s+(BLUEWUD[\s\S]+?)(?=\s*₹|\s*HSN|\s*\d+%)',
            r'(BLUEWUD[^\n]+?)(?=\s*[A-Z][\d]|\s*HSN|\s*₹)',
        ]

        for pattern in description_patterns:
            try:
                matches = re.findall(pattern, text, re.IGNORECASE | re.DOTALL)
                if matches:
                    best_match = max(matches, key=len) if isinstance(matches[0], str) else str(matches[0])
                    cleaned_description = self.clean_description(best_match)
                    if len(cleaned_description) > 20 and 'BLUEWUD' in cleaned_description.upper():
                        return cleaned_description
            except Exception:
                continue
        return ""

    def clean_description(self, description):
        """Clean the extracted description to match Excel format"""
        if not description:
            return ""

        description = re.sub(r'\s+', ' ', description).strip()
        
        unwanted_patterns = [
            r'HSN[\s:]*[\d]+',
            r'₹[\d,]+\.?\d*',
            r'\d+%\s*(CGST|SGST|IGST)',
            r'Sl\.?\s*No\.?',
            r'Unit\s*Price',
            r'Qty\s*Net',
            r'Tax\s*Rate',
            r'Tax\s*Type',
            r'Tax\s*Amount',
            r'Total\s*Amount',
            r'BLUEWUD CONCEPTS PRIVATE LIMITED[:\s\S]*?Description\s*Amount\s*1',
            r'Authorized Signatory[\s\S]*?Description\s*Amount\s*1',
            r'Order Number:.*?Invoice Date\s*:\s*[\d./]+\s*Description\s*Amount\s*1',
            r'BLUEWUD CONCEPTS PRIVATE LIMITED[:\s\S]*?Order Number:.*?Description\s*Amount\s*1',
        ]

        for pattern in unwanted_patterns:
            description = re.sub(pattern, '', description, flags=re.IGNORECASE)

        description = re.sub(r'\s+', ' ', description).strip()
        description = re.sub(r'^[\s\-|,.:;]+', '', description)
        description = re.sub(r'[\s\-|,.:;]+$', '', description)
        
        return description

    def extract_field(self, text, field_name):
        """Extract a single field using proven patterns"""
        if field_name == 'Description':
            return self.extract_description_from_table(text)

        patterns = self.extraction_patterns.get(field_name, [])
        
        for pattern in patterns:
            try:
                match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                if match:
                    value = match.group(1).strip()
                    if field_name == 'Customer Address':
                        value = self.clean_address(value)
                    elif field_name == 'Total Amount':
                        # STRICT: Only accept amounts ending in .00
                        if value.endswith('.00'):
                            return f"₹{value}"
                    
                    if value and field_name != 'Total Amount':
                        return value
            except Exception:
                continue

        # Fallbacks
        if field_name == 'Invoice Details':
            m = re.search(r'\b(UP-143350511-\d+|KA-BLX1-[\d-]+)\b', text)
            if m:
                return m.group(1)

        if field_name == 'Total Amount':
            # ONLY accept final amounts ending in .00
            patterns_strict = [
                r'Invoice Value[:\s]*₹?\s*([\d,]+\.00)',  # Most reliable
                r'₹\s*([\d,]{4,}\.00)(?![^\n]*(?:CGST|SGST|IGST|Tax))',  # Large amounts only
            ]

            for pattern in patterns_strict:
                matches = re.findall(pattern, text, re.IGNORECASE)
                if matches:
                    # Take the largest .00 amount
                    largest = max(matches, key=lambda x: float(x.replace(',', '')))
                    return f"₹{largest}"

        return ""

    def clean_address(self, address):
        """Clean up extracted address"""
        address = re.sub(r'\s+', ' ', address)
        address = re.sub(r'\n+', ' ', address)
        address = re.sub(r'State/UT Code.*', '', address)
        address = re.sub(r'Place of.*', '', address)
        return address.strip()

    def extract_invoice_data_from_bytes(self, pdf_bytes, filename):
        """Extract all fields from a single invoice PDF bytes"""
        text = self.extract_text_from_pdf_bytes(pdf_bytes)
        if not text.strip():
            return None

        extracted_data = {}
        field_names = ['Order Number', 'Order Date', 'Invoice Number',
                      'Invoice Details', 'Customer Address', 'Description', 'Total Amount']

        for field_name in field_names:
            extracted_data[field_name] = self.extract_field(text, field_name)

        # Add filename for reference
        extracted_data['Source File'] = filename
        
        return extracted_data

    def process_uploaded_files(self, uploaded_files):
        """Process uploaded PDF files and return results"""
        results = []
        successful = 0
        failed = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, uploaded_file in enumerate(uploaded_files):
            try:
                status_text.text(f"Processing: {uploaded_file.name}")
                
                pdf_bytes = uploaded_file.read()
                extracted_data = self.extract_invoice_data_from_bytes(pdf_bytes, uploaded_file.name)
                
                if extracted_data and any(str(val).strip() for val in extracted_data.values() if val):
                    results.append(extracted_data)
                    successful += 1
                else:
                    failed += 1
                    st.warning(f"No data extracted from {uploaded_file.name}")
                    
            except Exception as e:
                failed += 1
                st.error(f"Error processing {uploaded_file.name}: {e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        status_text.text(f"Processing complete! ✅ {successful} successful, ❌ {failed} failed")
        
        return results, successful, failed

def create_excel_file(results):
    """Create Excel file from results and return BytesIO object"""
    if not results:
        return None
    
    df = pd.DataFrame(results)
    
    # Ensure required columns exist
    required_columns = [
        'Order Number', 'Order Date', 'Invoice Number',
        'Customer Address', 'Invoice Details', 'Description', 'Total Amount', 'Source File'
    ]
    
    for col in required_columns:
        if col not in df.columns:
            df[col] = ""
    
    df = df[required_columns]
    
    # Create Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='invoice_data_extracted')
        
        # Set column widths
        worksheet = writer.sheets['invoice_data_extracted']
        column_widths = {
            'A': 20, 'B': 12, 'C': 15, 'D': 50, 'E': 25, 'F': 80, 'G': 12, 'H': 20
        }
        
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
    
    output.seek(0)
    return output

def main():
    st.set_page_config(
        page_title="Amazon Invoice Converter",
        page_icon="📊",
        layout="centered"
    )
    
    # Header
    st.markdown("""
    <div style="background-color: #2196F3; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
        <h1 style="color: white; text-align: center; margin: 0;">
            📊 Amazon Invoices PDF to Excel Converter
        </h1>
        <p style="color: white; text-align: center; margin: 10px 0 0 0;">
            Extract invoice data from Amazon PDF invoices into organized Excel format
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Instructions
    st.markdown("""
    ### How to use:
    1. **Upload PDF files**: Select one or multiple Amazon invoice PDF files
    2. **Process**: Click the process button to extract data
    3. **Download**: Get your organized Excel file with all invoice data
    """)
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose Amazon Invoice PDF files",
        type="pdf",
        accept_multiple_files=True,
        help="You can upload multiple PDF files at once"
    )
    
    if uploaded_files:
        st.success(f"📁 {len(uploaded_files)} PDF file(s) uploaded successfully!")
        
        # Show uploaded files
        with st.expander("View uploaded files"):
            for file in uploaded_files:
                st.write(f"• {file.name} ({file.size:,} bytes)")
        
        # Process button
        if st.button("🚀 Start Processing", type="primary"):
            with st.spinner("Processing PDF files..."):
                
                extractor = FinalPerfectExtractor()
                results, successful, failed = extractor.process_uploaded_files(uploaded_files)
                
                if results:
                    st.success("✅ Processing completed!")
                    
                    # Show results summary
                    total = successful + failed
                    success_rate = (successful / total * 100) if total > 0 else 0
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Successfully Processed", successful)
                    with col2:
                        st.metric("Failed", failed)
                    with col3:
                        st.metric("Success Rate", f"{success_rate:.1f}%")
                    
                    # Show extracted data preview
                    st.subheader("📋 Extracted Data Preview")
                    df_preview = pd.DataFrame(results)
                    st.dataframe(df_preview, width="stretch")git add -A
git commit -m "Force redeployment: update app with latest files"

                    
                    # Create and offer Excel download
                    excel_file = create_excel_file(results)
                    if excel_file:
                        st.download_button(
                            label="💾 Download Excel File",
                            data=excel_file,
                            file_name=f"amazon_invoices_extracted_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.info("📊 Your Amazon invoice data has been successfully extracted and organized!")
                    
                else:
                    st.error("❌ No data could be extracted from the uploaded files. Please check that the PDFs are valid Amazon invoices.")
    
    else:
        st.info("👆 Upload your Amazon invoice PDF files to get started")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; font-size: 0.8em;">
        Amazon Invoice Converter | Streamlit Version | Extract • Organize • Excel
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()