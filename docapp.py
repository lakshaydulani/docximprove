import streamlit as st
from docx import Document

st.title("Docx File Parser")

uploaded_file = st.file_uploader("Upload your file")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    for para in doc.paragraphs:
        st.write(para.text)

