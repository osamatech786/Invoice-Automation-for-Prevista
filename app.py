import streamlit as st
from docx import Document
import re
import time

# Set page configuration with a favicon
st.set_page_config(
    page_title="Prevista Invoice Automation",
    page_icon="https://lirp.cdn-website.com/d8120025/dms3rep/multi/opt/social-image-88w.png", 
    layout="centered"  # "centered" or "wide"
)

##############################
# Functions
##############################

def download_invoice_template():
    with open("resources/invoice_template.docx", "rb") as file:
        st.download_button(
            label="Download invoice template",
            data=file,
            file_name="invoice_template.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

def extract_total_from_invoice(docx_file):
    """
    Extract total amount from a DOCX invoice file
    """
    doc = Document(docx_file)
    total = 0
    
    # Search for total amount in document paragraphs
    for paragraph in doc.paragraphs:
        text = paragraph.text.lower()
        if 'total' in text:
            # Extract numbers from text
            numbers = re.findall(r'\d+\.?\d*', text)
            if numbers:
                # Take the last number found as total
                total = float(numbers[-1])
    return total

##############################
# Main
##############################

def main():
    # Initialize session state variables
    if "invoice_uploaded" not in st.session_state:
        st.session_state["invoice_uploaded"] = False
    if "processed" not in st.session_state:
        st.session_state["processed"] = False
    if "submitting" not in st.session_state:
        st.session_state["submitting"] = False

    # Logo space
    st.markdown(
        """
        <style>
            .logo {
                display: flex;
                justify-content: center;
                align-items: center;
                padding: 20px;
                border-bottom: 2px solid #f0f0f0;
                margin-bottom: 20px;
            }
        </style>
        """, unsafe_allow_html=True
    )

    st.image("resources/logo_removed_bg - enlarged.png", use_column_width=True)

    # Page title
    st.markdown(
        """
        <h2 style="text-align:center; color:#4CAF50;">Invoice Submission System</h2>
        """, unsafe_allow_html=True
    )

    # Download Invoice Template
    st.write("")
    st.divider()

    st.markdown(
        """
        <div style="text-align:left;">
            <p>Need an invoice template? Click below to download:</p>
        </div>
        """, unsafe_allow_html=True
    )
    download_invoice_template()
    st.divider()

    # File Upload Section
    st.markdown("<h3>Upload your Files</h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        invoice_file = st.file_uploader(
            "Upload Invoice (DOCX file)", 
            type=["docx"], 
            help="Upload your invoice here (mandatory)"
        )
    with col2:
        receipt_files = st.file_uploader(
            "Upload Expense Receipts (Optional, multiple)", 
            type=["jpg", "png", "pdf", "docx"], 
            accept_multiple_files=True,
            help="Upload your receipts here (optional)"
        )

    # Update state if invoice file is uploaded
    if invoice_file:
        st.session_state["invoice_uploaded"] = True
        st.success(f"Uploaded Invoice: {invoice_file.name}")
    else:
        st.session_state["invoice_uploaded"] = False

    # Display Uploaded Files
    if receipt_files:
        for receipt in receipt_files:
            st.success(f"Uploaded Receipt: {receipt.name}")

    # Process Button and Total Display
    st.markdown("<h3>Process Your Invoice</h3>", unsafe_allow_html=True)
    if st.button("Process"):
        if not st.session_state["invoice_uploaded"]:
            st.error("Please upload an invoice before processing!")
        else:
            st.session_state["processed"] = True  
            with st.spinner("Processing..."):
                # time.sleep(2)  # Simulating processing time
                st.session_state["total"] = extract_total_from_invoice(invoice_file)
                st.success("Invoice processed successfully!")

                # Display Total (if processed)
                st.write(f"Total: {st.session_state.get('total', 0)}")
                st.write("If the above total is correct, click Submit to complete your invoice submission:")

    # Submit Button Logic
    if st.session_state["invoice_uploaded"] and st.session_state["processed"]:
        if st.session_state["submitting"]:
            with st.spinner("Submitting..."):
                # time.sleep(5)  # Simulating submission

                # All Logic here from sharepoint file upload to update master sheet & Folder creation if doesn't exist

                # ....

                # ####################

                st.session_state["submitting"] = False
                st.success("Invoice Submitted Successfully!")
        else:
            if st.button("Submit"):
                st.session_state["submitting"] = True
                st.experimental_rerun()  # Rerun immediately to start the spinner
    else:
        st.button("Submit", disabled=True)


if __name__ == "__main__":
    main()

# Dev : https://linkedin.com/in/osamatech786