import streamlit as st
import fitz

from google import genai
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from io import BytesIO


# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="Notes Generator",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ---------------------------------------------------------
# CUSTOM INTERFACE DESIGN
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(
            135deg,
            #fffaf7 0%,
            #fff4f6 48%,
            #f8f4ff 100%
        );
    }

    .block-container {
        max-width: 1100px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    header[data-testid="stHeader"] {
        background: transparent;
    }

    .upload-title {
        color: #5b4b73;
        font-size: 21px;
        font-weight: 700;
        margin-top: 10px;
        margin-bottom: 8px;
    }

    .upload-description {
        color: #746b80;
        font-size: 15px;
        margin-bottom: 12px;
    }

    div[data-testid="stFileUploader"] section {
        background: rgba(255, 255, 255, 0.88);
        border: 2px dashed #b9a7d8;
        border-radius: 20px;
        padding: 18px;
        box-shadow: 0 8px 24px rgba(91, 75, 115, 0.10);
    }

    div[data-testid="stFileUploader"] section:hover {
        border-color: #8f72c2;
        background: #fdfbff;
    }

    div[data-testid="stFileUploader"] button {
        background: linear-gradient(135deg, #6f5aa8, #9174c8);
        color: white;
        border: none;
        border-radius: 999px;
        padding: 10px 22px;
        font-weight: 700;
        box-shadow: 0 6px 16px rgba(111, 90, 168, 0.25);
        transition: all 0.2s ease;
    }

    div[data-testid="stFileUploader"] button:hover {
        background: linear-gradient(135deg, #604b98, #8063b7);
        color: white;
        border: none;
        transform: translateY(-1px);
    }

    div[data-testid="stFileUploader"] small {
        color: #8a8193;
    }

    .supported-files-card {
        background: rgba(255, 255, 255, 0.72);
        border: 1px solid #ded4ec;
        border-radius: 13px;
        padding: 10px 14px;
        color: #695d76;
        font-size: 14px;
        text-align: center;
        margin-top: 8px;
        margin-bottom: 18px;
    }

    div.stButton > button,
    div.stDownloadButton > button {
        width: 100%;
        border-radius: 13px;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ---------------------------------------------------------
# GEMINI CLIENT CONNECTION
# ---------------------------------------------------------
try:
    client = genai.Client(
        api_key=st.secrets["GEMINI_API_KEY"]
    )

except KeyError:
    st.error(
        "Gemini API key was not found. "
        "Please add GEMINI_API_KEY inside "
        ".streamlit/secrets.toml."
    )
    st.stop()


# ---------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------
if "generated_notes" not in st.session_state:
    st.session_state.generated_notes = ""

if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = ""


# ---------------------------------------------------------
# PDF TEXT EXTRACTION
# ---------------------------------------------------------
def extract_pdf_text(uploaded_file):

    pdf_document = fitz.open(
        stream=uploaded_file.getvalue(),
        filetype="pdf"
    )

    text_blocks = []
    total_pages = len(pdf_document)

    for page in pdf_document:

        blocks = page.get_text("blocks")

        for block in blocks:

            block_text = block[4].strip()

            if block_text:
                text_blocks.append(block_text)

    pdf_document.close()

    extracted_text = "\n\n".join(text_blocks)

    return extracted_text, total_pages


# ---------------------------------------------------------
# WORD (.docx) TEXT EXTRACTION
# ---------------------------------------------------------
def extract_docx_text(uploaded_file):

    document = Document(
        BytesIO(uploaded_file.getvalue())
    )

    text_parts = []

    # Read normal paragraphs
    for paragraph in document.paragraphs:

        paragraph_text = paragraph.text.strip()

        if paragraph_text:
            text_parts.append(paragraph_text)

    # Read Word tables
    for table in document.tables:

        for row in table.rows:

            row_text = []

            for cell in row.cells:

                cell_text = cell.text.strip()

                if cell_text:
                    row_text.append(cell_text)

            if row_text:
                text_parts.append(" | ".join(row_text))

    return "\n\n".join(text_parts)


# ---------------------------------------------------------
# TXT TEXT EXTRACTION
# ---------------------------------------------------------
def extract_txt_text(uploaded_file):

    file_bytes = uploaded_file.getvalue()

    try:
        extracted_text = file_bytes.decode("utf-8")

    except UnicodeDecodeError:
        extracted_text = file_bytes.decode(
            "latin-1",
            errors="replace"
        )

    return extracted_text


# ---------------------------------------------------------
# POWERPOINT (.pptx) TEXT EXTRACTION
# ---------------------------------------------------------
def extract_pptx_text(uploaded_file):

    presentation = Presentation(
        BytesIO(uploaded_file.getvalue())
    )

    text_parts = []

    for slide_number, slide in enumerate(
        presentation.slides,
        start=1
    ):

        slide_text = []

        for shape in slide.shapes:

            # Read normal text boxes and titles
            if hasattr(shape, "text"):

                shape_text = shape.text.strip()

                if shape_text:
                    slide_text.append(shape_text)

            # Read tables inside PowerPoint
            if shape.has_table:

                for row in shape.table.rows:

                    row_text = []

                    for cell in row.cells:

                        cell_text = cell.text.strip()

                        if cell_text:
                            row_text.append(cell_text)

                    if row_text:
                        slide_text.append(
                            " | ".join(row_text)
                        )

        if slide_text:

            text_parts.append(
                f"Slide {slide_number}\n"
                + "\n".join(slide_text)
            )

    total_slides = len(presentation.slides)

    return "\n\n".join(text_parts), total_slides


# ---------------------------------------------------------
# EXCEL (.xlsx) TEXT EXTRACTION
# ---------------------------------------------------------
def extract_xlsx_text(uploaded_file):

    workbook = load_workbook(
        filename=BytesIO(uploaded_file.getvalue()),
        data_only=True,
        read_only=True
    )

    text_parts = []
    total_sheets = len(workbook.sheetnames)

    for worksheet in workbook.worksheets:

        sheet_data = []

        sheet_data.append(
            f"Worksheet: {worksheet.title}"
        )

        for row in worksheet.iter_rows(
            values_only=True
        ):

            row_values = []

            for cell_value in row:

                if cell_value is not None:

                    value = str(cell_value).strip()

                    if value:
                        row_values.append(value)

            if row_values:
                sheet_data.append(
                    " | ".join(row_values)
                )

        if len(sheet_data) > 1:
            text_parts.append(
                "\n".join(sheet_data)
            )

    workbook.close()

    return "\n\n".join(text_parts), total_sheets


# ---------------------------------------------------------
# CLEAN EXTRACTED TEXT
# ---------------------------------------------------------
def clean_extracted_text(text):

    cleaned_text = "\n\n".join(
        " ".join(paragraph.split())
        for paragraph in text.split("\n\n")
        if paragraph.strip()
    )

    return cleaned_text


# ---------------------------------------------------------
# GEMINI PROMPT
# ---------------------------------------------------------
def create_notes_prompt(study_material):

    prompt = f"""
You are an expert college study-notes creator.

Read only the study material provided below.
Do not add unrelated information.
Use clear, simple, grammatically correct English.
Make the output useful for college examinations.

Generate the response using exactly this structure:

# 📘 Chapter / Topic Title
Identify the most suitable title from the material.

## 📖 Definition
Give a clear and complete definition of the main topic.
Use one short paragraph.

## ⭐ Important Points
Provide 8 to 15 important points.
Use simple bullet points.
Do not use checkmark symbols.

## 📝 Detailed Summary
Explain the complete topic in simple English.
Use short paragraphs and preserve important concepts,
terms, formulas, or steps found in the material.

## ❓ 2-Mark Questions
Generate 5 short-answer questions.

## ❓ 5-Mark Questions
Generate 5 medium-answer questions.

## ❓ 10-Mark Questions
Generate 3 long-answer questions.

## 🧠 Multiple Choice Questions
Generate 5 MCQs.

Each MCQ must contain four options:
A, B, C, and D.

After every MCQ, write:
**Answer:** followed by the correct option and answer.

## 🎯 Quick Quiz
Generate 5 questions for student self-practice.
Do not provide answers in this section.

Important rules:

- Base the output only on the supplied material.
- Do not invent facts that are absent from the material.
- Avoid repeated questions and repeated points.
- Use professional Markdown headings.
- Keep the language easy for a college student.

STUDY MATERIAL:

{study_material}
"""

    return prompt


# ---------------------------------------------------------
# APPLICATION HEADING
# ---------------------------------------------------------
st.markdown(
    """
    <h1 style="
        text-align: center;
        color: black;
        font-size: 52px;
        font-weight: 800;
        margin-bottom: 10px;
    ">
        📚 Notes Generator
    </h1>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <p style="
        text-align:center;
        font-size:24px;
        color:#7562A8;
        font-weight:700;
        margin-bottom:5px;
    ">
        Smart Study Assistant
    </p>

    <p style="
        text-align:center;
        font-size:17px;
        color:#9A8CBC;
        margin-bottom:25px;
    ">
        Turn your study files into clear AI notes.
    </p>
    """,
    unsafe_allow_html=True
)


# ---------------------------------------------------------
# FILE UPLOADER
# ---------------------------------------------------------
st.markdown(
    """
    <div class="upload-title">📂 Upload Study Material</div>
    <div class="upload-description">
        Choose your college file and let AI prepare structured study notes.
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "Upload Study Material",
    type=[
        "pdf",
        "docx",
        "txt",
        "pptx",
        "xlsx"
    ],
    key="study_material_uploader",
    label_visibility="collapsed"
)

st.markdown(
    """
    <div class="supported-files-card">
        Supported formats: <b>PDF • DOCX • TXT • PPTX • XLSX</b>
    </div>
    """,
    unsafe_allow_html=True
)


# ---------------------------------------------------------
# FILE PROCESSING
# ---------------------------------------------------------
if uploaded_file is not None:

    # Clear old notes when a new file is uploaded
    if (
        st.session_state.uploaded_file_name
        != uploaded_file.name
    ):
        st.session_state.generated_notes = ""
        st.session_state.uploaded_file_name = (
            uploaded_file.name
        )

    st.success(
        f"{uploaded_file.name} uploaded successfully!"
    )

    file_extension = (
        uploaded_file.name
        .split(".")[-1]
        .lower()
    )

    all_text = ""
    file_processed = False


    # -----------------------------------------------------
    # PDF PROCESSING
    # -----------------------------------------------------
    if file_extension == "pdf":

        try:
            all_text, total_pages = extract_pdf_text(
                uploaded_file
            )

            st.success(
                "PDF opened successfully!"
            )

            file_processed = True

        except Exception as error:

            st.error(
                "The PDF could not be processed."
            )

            st.code(str(error))


    # -----------------------------------------------------
    # WORD PROCESSING
    # -----------------------------------------------------
    elif file_extension == "docx":

        try:
            all_text = extract_docx_text(
                uploaded_file
            )

            st.success(
                "Word document opened successfully!"
            )

            file_processed = True

        except Exception as error:

            st.error(
                "The Word document could not be processed."
            )

            st.code(str(error))


    # -----------------------------------------------------
    # TXT PROCESSING
    # -----------------------------------------------------
    elif file_extension == "txt":

        try:
            all_text = extract_txt_text(
                uploaded_file
            )

            st.success(
                "TXT file opened successfully!"
            )

            file_processed = True

        except Exception as error:

            st.error(
                "The TXT file could not be processed."
            )

            st.code(str(error))


    # -----------------------------------------------------
    # POWERPOINT PROCESSING
    # -----------------------------------------------------
    elif file_extension == "pptx":

        try:
            all_text, total_slides = extract_pptx_text(
                uploaded_file
            )

            st.success(
                "PowerPoint presentation opened successfully!"
            )

            file_processed = True

        except Exception as error:

            st.error(
                "The PowerPoint file could not be processed."
            )

            st.code(str(error))


    # -----------------------------------------------------
    # EXCEL PROCESSING
    # -----------------------------------------------------
    elif file_extension == "xlsx":

        try:
            all_text, total_sheets = extract_xlsx_text(
                uploaded_file
            )

            st.success(
                "Excel workbook opened successfully!"
            )

            file_processed = True

        except Exception as error:

            st.error(
                "The Excel file could not be processed."
            )

            st.code(str(error))


    # -----------------------------------------------------
    # COMMON TEXT PREVIEW AND AI GENERATION
    # -----------------------------------------------------
    if file_processed:

        if all_text.strip():

            clean_text = clean_extracted_text(
                all_text
            )

            maximum_characters = 50000

            text_for_ai = clean_text[
                :maximum_characters
            ]

            if len(clean_text) > maximum_characters:

                st.warning(
                    "This file contains a large amount of text. "
                    "For this version, the first 50,000 "
                    "characters will be used."
                )


            # -------------------------------------------------
            # GENERATE AI NOTES BUTTON
            # -------------------------------------------------
            if st.button(
                "✨ Generate AI Study Notes",
                type="primary",
                key="generate_ai_notes_button"
            ):

                prompt = create_notes_prompt(
                    text_for_ai
                )

                try:

                    with st.spinner(
                        "Gemini AI is reading the material "
                        "and preparing your study notes..."
                    ):

                        response = (
                            client.models.generate_content(
                                model="gemini-3.5-flash",
                                contents=prompt
                            )
                        )

                    if response.text:

                        st.session_state.generated_notes = (
                            response.text
                        )

                        st.success(
                            "AI study notes generated successfully!"
                        )

                    else:

                        st.warning(
                            "Gemini did not return any notes. "
                            "Please try again."
                        )

                except Exception as error:

                    st.error(
                        "AI notes generation failed."
                    )

                    st.code(str(error))


            # -------------------------------------------------
            # DISPLAY GENERATED NOTES
            # -------------------------------------------------
            if st.session_state.generated_notes:

                st.divider()

                st.subheader(
                    "📝 Generated Study Notes"
                )

                st.markdown(
                    st.session_state.generated_notes
                )

                st.download_button(
                    label="⬇️ Download Notes as Text File",
                    data=st.session_state.generated_notes,
                    file_name="AI_Generated_Study_Notes.txt",
                    mime="text/plain",
                    key="download_notes_button"
                )

        else:

            st.warning(
                "No readable text was found in this file."
            )

else:

    st.caption(
        "Choose a file above to create clear AI-powered study notes."
    )