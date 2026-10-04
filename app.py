import streamlit as st
import re
import fitz

from google import genai
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from chatbot import render_chatbot
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
# GOOGLE SEARCH CONSOLE VERIFICATION
# ---------------------------------------------------------
st.html(
    """
    <script>
    const meta = document.createElement("meta");
    meta.name = "google-site-verification";
    meta.content = "Ejh0TxxtsuIWMbqQBvTuxOGu7OTL2mnxDqn0XPXg8tY";
    document.head.appendChild(meta);
    </script>
    """,
    unsafe_allow_javascript=True
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

    /* ---------------------------------------------------------
       MODERN CHATBOT DESIGN
       --------------------------------------------------------- */

    .chatbot-card {
        background: rgba(255, 255, 255, 0.88);
        border: 1px solid #e3d8f0;
        border-radius: 24px;
        padding: 24px 26px 18px 26px;
        margin-top: 8px;
        margin-bottom: 22px;
        box-shadow: 0 12px 30px rgba(91, 75, 115, 0.10);
    }

    .chatbot-card-title {
        color: #4f3f82;
        font-size: 25px;
        font-weight: 800;
        margin-bottom: 4px;
    }

    .chatbot-card-subtitle {
        color: #827791;
        font-size: 14px;
        margin-bottom: 18px;
    }

    div[data-testid="stChatMessage"]:has(
        div[data-testid="stChatMessageAvatarIcon-assistant"]
    ) {
        background: #f7f3ff;
        border: 1px solid #e5dcf3;
        border-radius: 18px;
        padding: 10px 14px;
        margin: 10px 18% 12px 0;
        box-shadow: 0 4px 12px rgba(91, 75, 115, 0.06);
    }

    div[data-testid="stChatMessage"]:has(
        div[data-testid="stChatMessageAvatarIcon-user"]
    ) {
        background: linear-gradient(
            135deg,
            #eadfff,
            #f7e9f5
        );
        border: 1px solid #d9c9ee;
        border-radius: 18px;
        padding: 10px 14px;
        margin: 10px 0 12px 18%;
        box-shadow: 0 4px 12px rgba(111, 90, 168, 0.08);
    }

    div[data-testid="stChatInput"] {
        background: #ffffff;
        border: 1px solid #d8cce9;
        border-radius: 18px;
        box-shadow: 0 5px 16px rgba(91, 75, 115, 0.08);
    }

    div[data-testid="stChatInput"] textarea {
        font-size: 15px;
    }

    /* FINAL CHATBOT PANEL */
    .chatbot-panel { background:rgba(255,255,255,0.94); border:1px solid #dfd4ee; border-radius:26px; padding:24px; margin:8px 0 24px; box-shadow:0 14px 32px rgba(91,75,115,0.10); }
    .chatbot-panel-title { color:#4f3f82; font-size:25px; font-weight:800; margin-bottom:4px; }
    .chatbot-panel-subtitle { color:#81768e; font-size:14px; margin-bottom:18px; }
    .chat-scroll-area { background:linear-gradient(180deg,#fcfaff,#fffafd); border:1px solid #ece4f4; border-radius:20px; padding:16px; min-height:90px; max-height:520px; overflow-y:auto; }
    .chat-row-user { display:flex; justify-content:flex-end; margin:10px 0; }
    .chat-row-ai { display:flex; justify-content:flex-start; margin:10px 0; }
    .chat-bubble-user { max-width:72%; background:linear-gradient(135deg,#eadfff,#f7e8f5); border:1px solid #d9c9ee; border-radius:18px 18px 5px 18px; padding:11px 15px; color:#44365f; line-height:1.5; }
    .chat-bubble-ai { max-width:78%; background:#ffffff; border:1px solid #e3d9ef; border-radius:18px 18px 18px 5px; padding:11px 15px; color:#4f4758; line-height:1.55; box-shadow:0 4px 12px rgba(91,75,115,0.05); }
    .chat-avatar { font-size:18px; margin:0 7px; align-self:flex-end; }
    .chat-empty { color:#9a91a3; text-align:center; padding:34px 10px; font-size:14px; }
    div[data-testid="stForm"] { border:0 !important; padding:0 !important; }
    .chat-input-label { color:#6d6178; font-size:13px; font-weight:700; margin:14px 0 6px 2px; }

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

if "document_text" not in st.session_state:
    st.session_state.document_text = ""

if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []

if "generated_5_mark_questions" not in st.session_state:
    st.session_state.generated_5_mark_questions = ""

if "generated_2_mark_questions" not in st.session_state:
    st.session_state.generated_2_mark_questions = ""

if "generated_10_mark_questions" not in st.session_state:
    st.session_state.generated_10_mark_questions = ""

if "generated_mcqs" not in st.session_state:
    st.session_state.generated_mcqs = ""

if "generated_quiz" not in st.session_state:
    st.session_state.generated_quiz = ""
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

## 📌 Additional Exam Preparation
Additional exam questions, MCQs, and quiz will be generated
separately when the student requests them.

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
# 5-MARK QUESTIONS PROMPT
# ---------------------------------------------------------
def create_5_mark_prompt(study_material):

    prompt = f"""
You are an expert college examination question creator.

Read only the study material provided below.
Generate exactly 5 important 5-mark examination questions
based only on the supplied material.

Rules:
- Questions must be suitable for a college examination.
- Questions should test understanding, explanation, comparison,
  steps, concepts, or applications found in the material.
- Do not add unrelated information.
- Do not repeat questions.
- Use simple, clear English.
- Do not generate 2-mark questions, 10-mark questions, MCQs,
  or quiz questions.
- Output only the 5-mark questions.

STUDY MATERIAL:

{study_material}
"""

    return prompt


def create_2_mark_prompt(study_material):
    return f"""
You are an expert college examination question creator.
Read only the study material provided below.
Generate exactly 5 important 2-mark examination questions.

Rules:
- Questions must be short and suitable for a college examination.
- Focus on definitions, terms, facts, concepts, or short explanations.
- Do not add unrelated information or repeat questions.
- Use simple, clear English.
- Do not generate 5-mark, 10-mark, MCQ, or quiz questions.
- Output only the 2-mark questions.

STUDY MATERIAL:
{study_material}
"""


def create_10_mark_prompt(study_material):
    return f"""
You are an expert college examination question creator.
Read only the study material provided below.
Generate exactly 3 important 10-mark examination questions.

Rules:
- Questions must be suitable for a college examination.
- Focus on detailed explanations, comparisons, architecture, processes, steps, or applications found in the material.
- Do not add unrelated information or repeat questions.
- Use simple, clear English.
- Do not generate 2-mark, 5-mark, MCQ, or quiz questions.
- Output only the 10-mark questions.

STUDY MATERIAL:
{study_material}
"""


def create_mcq_prompt(study_material):
    return f"""
You are an expert college examination question creator.
Read only the study material provided below.
Generate exactly 5 multiple-choice questions.

Rules:
- Each question must have exactly four options: A, B, C, and D.
- Provide the correct answer after each question.
- Base everything only on the supplied material.
- Do not add unrelated information or repeat questions.
- Use simple, clear English.
- Do not generate 2-mark, 5-mark, or 10-mark questions.

STUDY MATERIAL:
{study_material}
"""


def create_quiz_prompt(study_material):
    return f"""
You are an expert college study assistant.
Read only the study material provided below.
Generate exactly 5 self-practice quiz questions.

Rules:
- Cover important concepts from the material.
- Use short-answer and understanding-based questions.
- Do not provide answers.
- Do not add unrelated information or repeat questions.
- Use simple, clear English.
- Do not generate 2-mark, 5-mark, 10-mark questions, or MCQs.

STUDY MATERIAL:
{study_material}
"""


# ---------------------------------------------------------
# CHATBOT PROMPT
# ---------------------------------------------------------
def split_generated_notes(notes_text):
    """
    Separate the single Gemini response into visual sections.

    The original generated text is kept unchanged in session state/downloads.
    This function is only used for displaying the result as separate cards.
    """
    section_patterns = [
        ("📖 Definition", r"##\s*📖\s*Definition"),
        ("⭐ Important Points", r"##\s*⭐\s*Important Points"),
        ("📝 Detailed Summary", r"##\s*📝\s*Detailed Summary"),
        ("❓ 2-Mark Questions", r"##\s*❓\s*2-Mark Questions"),
        ("❓ 5-Mark Questions", r"##\s*❓\s*5-Mark Questions"),
        ("❓ 10-Mark Questions", r"##\s*❓\s*10-Mark Questions"),
        ("🧠 Multiple Choice Questions", r"##\s*🧠\s*Multiple Choice Questions"),
        ("🎯 Quick Quiz", r"##\s*🎯\s*Quick Quiz"),
    ]

    sections = []

    # Find all known headings and their positions.
    matches = []
    for title, pattern in section_patterns:
        match = re.search(pattern, notes_text, flags=re.IGNORECASE)
        if match:
            matches.append((match.start(), match.end(), title))

    matches.sort(key=lambda item: item[0])

    # Keep the title/introduction before the first known section.
    if matches:
        intro = notes_text[:matches[0][0]].strip()
        if intro:
            sections.append(("📘 Chapter / Topic", intro))

        for index, (start_pos, end_pos, title) in enumerate(matches):
            next_start = (
                matches[index + 1][0]
                if index + 1 < len(matches)
                else len(notes_text)
            )

            content = notes_text[end_pos:next_start].strip()

            if content:
                sections.append((title, content))
    else:
        sections.append(("📝 Study Notes", notes_text.strip()))

    return sections


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
    # Clear old notes and chat when a new file is uploaded
    if (
        st.session_state.uploaded_file_name
        != uploaded_file.name
    ):
        st.session_state.generated_notes = ""
        st.session_state.generated_2_mark_questions = ""
        st.session_state.generated_5_mark_questions = ""
        st.session_state.generated_10_mark_questions = ""
        st.session_state.generated_mcqs = ""
        st.session_state.generated_quiz = ""
        st.session_state.document_text = ""
        st.session_state.chat_messages = []

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

            # Store the complete cleaned document for the chatbot
            st.session_state.document_text = clean_text

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
                                model="gemini-3.5-flash-lite",
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
            # OPTIONAL EXAM PREPARATION BUTTONS
            # -------------------------------------------------
            if st.session_state.generated_notes:

                exam_cols = st.columns(5)

                exam_options = [
                    ("📝 2-Mark", "2_mark", create_2_mark_prompt, "generated_2_mark_questions", "2-mark questions"),
                    ("📚 5-Mark", "5_mark", create_5_mark_prompt, "generated_5_mark_questions", "5-mark questions"),
                    ("📖 10-Mark", "10_mark", create_10_mark_prompt, "generated_10_mark_questions", "10-mark questions"),
                    ("🧠 MCQs", "mcq", create_mcq_prompt, "generated_mcqs", "MCQs"),
                    ("🎯 Quiz", "quiz", create_quiz_prompt, "generated_quiz", "quiz questions"),
                ]

                for column, (label, button_key, prompt_function, state_key, result_name) in zip(exam_cols, exam_options):
                    with column:
                        if st.button(label, key=f"generate_{button_key}_button"):
                            exam_prompt = prompt_function(text_for_ai)
                            try:
                                with st.spinner(f"Gemini AI is preparing {result_name}..."):
                                    exam_response = client.models.generate_content(
                                        model="gemini-3.5-flash-lite",
                                        contents=exam_prompt
                                    )

                                if exam_response.text:
                                    st.session_state[state_key] = exam_response.text.strip()
                                    st.success(f"{result_name.capitalize()} generated successfully!")
                                else:
                                    st.warning(f"Gemini did not return any {result_name}. Please try again.")
                            except Exception as error:
                                st.error(f"{result_name.capitalize()} generation failed.")
                                st.code(str(error))

                optional_sections = [
                    ("❓ 2-Mark Questions", "generated_2_mark_questions"),
                    ("❓ 5-Mark Questions", "generated_5_mark_questions"),
                    ("❓ 10-Mark Questions", "generated_10_mark_questions"),
                    ("🧠 Multiple Choice Questions", "generated_mcqs"),
                    ("🎯 Quick Quiz", "generated_quiz"),
                ]

                for section_title, state_key in optional_sections:
                    section_content = st.session_state.get(state_key, "")
                    if section_content:
                        st.markdown(
                            f"""
                            <div style="background:linear-gradient(135deg,#fff7f2,#fff1f5);border:2px solid #f2b5c8;border-radius:18px;padding:18px 22px;margin:16px 0 10px 0;box-shadow:0 6px 18px rgba(120,80,120,0.08);">
                                <div style="color:#8b3f63;font-size:22px;font-weight:800;margin-bottom:10px;">
                                    {section_title}
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                        st.markdown(section_content)

            # -------------------------------------------------
            # CHATBOT
            # -------------------------------------------------
            render_chatbot(
                client,
                st.session_state.document_text
            )

            # -------------------------------------------------
            # DISPLAY GENERATED NOTES AS SEPARATE STUDY SECTIONS
            # -------------------------------------------------
            if st.session_state.generated_notes:

                st.divider()

                notes_col, download_col = st.columns(
                    [0.72, 0.28]
                )

                with notes_col:
                    st.subheader("📝 Generated Study Notes")

                with download_col:
                    st.download_button(
                        label="⬇️ Download Notes",
                        data=st.session_state.generated_notes,
                        file_name="AI_Generated_Study_Notes.txt",
                        mime="text/plain",
                        key="download_notes_button"
                    )

                # Intro / topic title
                sections = split_generated_notes(
                    st.session_state.generated_notes
                )

                for section_title, section_content in sections:

                    # Keep important exam questions visually separate.
                    if section_title in [
                        "❓ 2-Mark Questions",
                        "❓ 5-Mark Questions",
                        "❓ 10-Mark Questions"
                    ]:
                        st.markdown(
                            f"""
                            <div style="
                                background: linear-gradient(
                                    135deg,
                                    #fff7f2,
                                    #fff1f5
                                );
                                border: 2px solid #f2b5c8;
                                border-radius: 18px;
                                padding: 18px 22px;
                                margin: 16px 0;
                                box-shadow: 0 6px 18px rgba(
                                    120, 80, 120, 0.08
                                );
                            ">
                                <div style="
                                    color: #8b3f63;
                                    font-size: 22px;
                                    font-weight: 800;
                                    margin-bottom: 10px;
                                ">
                                    {section_title}
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                        st.markdown(section_content)

                    else:
                        st.markdown(
                            f"""
                            <div style="
                                background: rgba(255,255,255,0.88);
                                border: 1px solid #ded4ec;
                                border-radius: 16px;
                                padding: 16px 20px 6px 20px;
                                margin: 14px 0;
                                box-shadow: 0 5px 16px rgba(
                                    91, 75, 115, 0.07
                                );
                            ">
                                <div style="
                                    color: #4f3f82;
                                    font-size: 21px;
                                    font-weight: 800;
                                    margin-bottom: 8px;
                                ">
                                    {section_title}
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                        st.markdown(section_content)


        else:

            st.warning(
                "No readable text was found in this file."
            )

else:

    st.caption(
        "Choose a file above to create clear AI-powered study notes."
    )
