import streamlit as st
from openai import OpenAI
from docx import Document
from io import BytesIO
from docx.shared import Pt


def modify_docx_styles(doc, output_path):
    
    for para in doc.paragraphs:
        
        # Check if the paragraph style is a heading style
        if para.style.name.startswith('Heading'):
            # Set font size for headings
            for run in para.runs:
                run.font.size = Pt(36)  # 36 points in half-points (36 * 20)
        else:
             for run in para.runs:
                run.font.size = Pt(14)
                
        # Set the font style for all text
        for run in para.runs:
            run.font.name = "EYInterstate Light"
            # run.font.size = Pt(14)
    
    # Save the modified document
    doc.save(output_path)


def load_docx(file):
    """Load DOCX file into a Document object."""
    return Document(file)

def docx_to_text(doc):
    """Extract text from DOCX Document."""
    return '\n\n'.join([para.text for para in doc.paragraphs])

def text_to_docx(text):
    """Convert text back into a DOCX Document."""
    doc = Document()
    for line in text.split('\n'):
        doc.add_paragraph(line)
    return doc

def improve_text_with_openai(text, api_key):
    
    client = OpenAI(
    api_key = api_key,
    )
    response = client.completions.create(
  model = "gpt-3.5-turbo-instruct",
  prompt = text,
  max_tokens = 7,
  temperature = 0
)
    return response.choices[0].text.strip()

def save_docx(doc):
    """Save Document object to a BytesIO object and return it."""
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def main():
    st.title("DOCX Enhancer")

    uploaded_file = st.file_uploader("Upload your file", type=['docx'])
    api_key = st.secrets["OPENAI_API_KEY"]

    if uploaded_file and api_key:
        with st.spinner('Working on your document...'):
            doc = load_docx(uploaded_file)
            # modify_docx_styles(doc,"modified_file.docx")
            # doc = load_docx("modified_file.docx")
            original_text = docx_to_text(doc)
            improved_text = improve_text_with_openai(original_text, api_key)
            new_doc = text_to_docx(improved_text)
            buffer = save_docx(new_doc)

            st.success("Enhancement Complete! Download your improved DOCX below.")
            st.download_button(label="Download DOCX", data=buffer, file_name="modified_file.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    main()
