import streamlit as st
from msdox2 import doc  # Import the fully prepared doc from msdox2.py
import os

st.set_page_config(page_title="Form A Generator")
st.title("Form A DOCX Generator")

if st.button("Generate Form A"):
    file_name = "FormA.docx"
    # Save the prepared document
    doc.save(file_name)

    # Provide download button
    with open(file_name, "rb") as f:
        st.download_button(
            label="Download Form A",
            data=f,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
