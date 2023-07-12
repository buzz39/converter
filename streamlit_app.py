#streamlit web app for converting word to pdf
import streamlit as st
import pdf2docx
import docx2pdf
import os
import base64
from PIL import Image
import io
import time
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go

st.title("PDF to Word Converter")
st.subheader("Convert PDF to Word with formatting")

st.sidebar.title("PDF to Word Converter")
st.sidebar.subheader("Convert PDF to Word with formatting")

st.sidebar.subheader("Select the file type to convert")
file_type = st.sidebar.selectbox("File Type", ["PDF to Word", "Word to PDF"])

#logic for pdf to word
if file_type == "PDF to Word":
    st.subheader("Convert PDF to Word with formatting")
    pdf_file = st.file_uploader("Upload PDF", type=['pdf'])
    if pdf_file is not None:
        st.write("PDF file uploaded successfully")
        st.write("Converting PDF to Word with formatting...")
        with st.spinner('Wait for it...'):
            time.sleep(5)
        st.success("Converted successfully! Download the Word file below.")
        word_file = pdf_file.name.replace(".pdf", ".docx")
        pdf2docx.parse(pdf_file, word_file)
        with open(word_file, "rb") as file:
            btn = st.download_button(
                label="Download",
                data=file,
                file_name=word_file,
                mime="application/octet-stream"
            )
        if btn:
            os.remove(pdf_file.name)
            os.remove(word_file)
    else:
        st.write("Please upload a PDF file to convert to Word.")

#logic for word to pdf
if file_type == "Word to PDF":
    st.subheader("Convert Word to PDF")
    word_file = st.file_uploader("Upload Word", type=['docx'])
    if word_file is not None:
        st.write("Word file uploaded successfully")
        st.write("Converting Word to PDF...")
        with st.spinner('Wait for it...'):
            time.sleep(5)
        st.success("Converted successfully! Download the PDF file below.")
        pdf_file = word_file.name.replace(".docx", ".pdf")
        docx2pdf.convert(word_file.name, pdf_file)
        with open(pdf_file, "rb") as file:
            btn = st.download_button(
                label="Download",
                data=file,
                file_name=pdf_file,
                mime="application/octet-stream"
            )
        if btn:
            os.remove(word_file.name)
            os.remove(pdf_file)
    else:
        st.write("Please upload a Word file to convert to PDF.")

st.sidebar.subheader("About App")

st.sidebar.info("This web app is made as part of the tutorial on converting PDF to Word with formatting.")

st.sidebar.subheader("By")
st.sidebar.text("Gagan Thakur")
