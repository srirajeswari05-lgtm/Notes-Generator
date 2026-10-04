import streamlit as st


# ---------------------------------------------------------
# CHATBOT PROMPT
# ---------------------------------------------------------
def create_chat_prompt(question, document_text, chat_history):
    history_text = ""

    for message in chat_history:
        role = message.get("role", "")
        content = message.get("content", "")

        if role == "user":
            history_text += f"Student: {content}\n"
        elif role == "assistant":
            history_text += f"Assistant: {content}\n"

    prompt = f"""
You are an AI college study assistant.

Your goal is to give helpful, accurate, and easy-to-understand answers.

Use the uploaded study material as the PRIMARY source whenever the
student's question is related to that material.

You may also use your general knowledge when it helps answer the
student's question, especially when:
- the question asks for a general explanation,
- the uploaded material does not contain enough information,
- the student asks for an example, clarification, or background,
- or the question is not directly related to the uploaded document.

IMPORTANT RULES:
- If the question is about the uploaded document, prioritize the
  information from that document.
- Do not contradict the uploaded material without clearly explaining
  the difference.
- If the uploaded material does not contain enough information,
  supplement the answer with general knowledge instead of simply
  refusing to answer.
- When you use information beyond the uploaded material, make it clear
  that it is additional/general information when that distinction matters.
- Do not invent information and do not present guesses as facts.
- Explain difficult concepts in simple English.
- Give clear answers suitable for a college student.
- Keep normal answers SHORT and focused: usually 3 to 6 sentences or a few bullet points.
- For a simple definition question such as "What is HTML?", give a concise definition,
  one or two key points, and stop.
- Do not give a long lecture unless the student asks for detail, examples, steps,
  or a full explanation.
- For "explain in detail" questions, explain step by step when useful.
- For "summarize" questions about the document, summarize the document content first
  and do not replace it with a generic summary.
- Use previous conversation to understand follow-up questions.
- Keep answers focused on exactly what the student asked.

PREVIOUS CONVERSATION:
{history_text}

UPLOADED STUDY MATERIAL:
{document_text}

STUDENT'S CURRENT QUESTION:
{question}

Now answer the student's question naturally.
Use the uploaded material when relevant and supplement it with general
knowledge when needed.
"""

    return prompt


# ---------------------------------------------------------
# CHATBOT DISPLAY HELPERS
# ---------------------------------------------------------
def format_answer_html(text):
    """Convert a few common Gemini Markdown patterns to safe HTML."""
    import html
    import re

    safe = html.escape(text)
    safe = safe.replace("\n", "<br>")
    safe = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", safe)
    safe = re.sub(r"`([^`]+)`", r"<code>\1</code>", safe)
    return safe


# ---------------------------------------------------------
# CHATBOT UI
# ---------------------------------------------------------
def render_chatbot(client, document_text):

    # Create chat history only once.
    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = []

    # -----------------------------------------------------
    # COMPACT CHATBOT DESIGN
    # -----------------------------------------------------
    st.markdown(
        """
        <style>
        .chatbot-panel {
            background: rgba(255,255,255,0.94);
            border: 1px solid #dfd4ee;
            border-radius: 20px;
            padding: 14px 20px 12px 20px;
            margin: 8px 0 14px 0;
            box-shadow: 0 8px 22px rgba(91,75,115,0.08);
        }

        .chatbot-panel-title {
            color: #4f3f82;
            font-size: 21px;
            font-weight: 800;
            margin: 0;
        }

        .chatbot-panel-subtitle {
            color: #81768e;
            font-size: 12px;
            margin: 3px 0 0 0;
        }

        .chat-scroll-area {
            background: linear-gradient(180deg,#fcfaff,#fffafd);
            border: 1px solid #ece4f4;
            border-radius: 18px;
            padding: 12px 14px;
            min-height: 70px;
            max-height: 420px;
            overflow-y: auto;
            margin-bottom: 8px;
        }

        .chat-row-user {
            display: flex;
            justify-content: flex-end;
            margin: 8px 0;
        }

        .chat-row-ai {
            display: flex;
            justify-content: flex-start;
            margin: 8px 0;
        }

        .chat-bubble-user {
            max-width: 72%;
            background: linear-gradient(135deg,#eadfff,#f7e8f5);
            border: 1px solid #d9c9ee;
            border-radius: 17px 17px 5px 17px;
            padding: 9px 13px;
            color: #44365f;
            line-height: 1.45;
            font-size: 14px;
        }

        .chat-bubble-ai {
            max-width: 78%;
            background: #ffffff;
            border: 1px solid #e3d9ef;
            border-radius: 17px 17px 17px 5px;
            padding: 9px 13px;
            color: #4f4758;
            line-height: 1.5;
            font-size: 14px;
            box-shadow: 0 3px 10px rgba(91,75,115,0.04);
        }

        .chat-avatar {
            font-size: 16px;
            margin: 0 6px;
            align-self: flex-end;
        }

        .chat-empty {
            color: #9a91a3;
            text-align: center;
            padding: 22px 10px;
            font-size: 13px;
        }

        div[data-testid="stForm"] {
            border: 0 !important;
            padding: 0 !important;
        }

        .chat-input-label {
            color: #6d6178;
            font-size: 12px;
            font-weight: 700;
            margin: 8px 0 4px 2px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # -----------------------------------------------------
    # COMPACT HEADER
    # -----------------------------------------------------
    st.markdown(
        """
        <div class="chatbot-panel">
            <div class="chatbot-panel-title">
                💬 Chat with Your Study Material
            </div>
            <div class="chatbot-panel-subtitle">
                Ask questions about your material or any general topic.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # -----------------------------------------------------
    # CHAT HISTORY
    # -----------------------------------------------------
    chat_html = '<div class="chat-scroll-area">'

    if not st.session_state.chat_messages:
        chat_html += (
            '<div class="chat-empty">'
            '👋 Hi! Ask me anything about your uploaded study material.'
            '</div>'
        )

    for message in st.session_state.chat_messages:
        content = message["content"]

        safe_content = format_answer_html(content)

        if message["role"] == "user":
            chat_html += f"""
            <div class="chat-row-user">
                <div class="chat-bubble-user">{safe_content}</div>
                <div class="chat-avatar">👤</div>
            </div>
            """
        else:
            chat_html += f"""
            <div class="chat-row-ai">
                <div class="chat-avatar">🤖</div>
                <div class="chat-bubble-ai">{safe_content}</div>
            </div>
            """

    chat_html += '</div>'
    st.html(chat_html)

    # -----------------------------------------------------
    # QUESTION INPUT
    # -----------------------------------------------------
    with st.form("study_chat_form", clear_on_submit=True):

        st.markdown(
            '<div class="chat-input-label">Ask your question</div>',
            unsafe_allow_html=True
        )

        input_col, send_col = st.columns([0.88, 0.12])

        with input_col:
            user_question = st.text_input(
                "Question",
                placeholder="e.g. What is HTML?",
                label_visibility="collapsed"
            )

        with send_col:
            send_question = st.form_submit_button("➤")

    # -----------------------------------------------------
    # GEMINI RESPONSE
    # -----------------------------------------------------
    if send_question and user_question.strip():

        user_question = user_question.strip()

        st.session_state.chat_messages.append(
            {"role": "user", "content": user_question}
        )

        with st.spinner("AI is thinking..."):
            try:
                chat_prompt = create_chat_prompt(
                    user_question,
                    document_text[:50000],
                    st.session_state.chat_messages[:-1]
                )

                chat_response = client.models.generate_content(
                    model="gemini-3.5-flash-lite",
                    contents=chat_prompt
                )

                if chat_response.text:
                    answer = chat_response.text.strip()
                else:
                    answer = (
                        "I could not generate an answer. "
                        "Please try again."
                    )

            except Exception:
                answer = (
                    "⚠️ The AI is temporarily busy. "
                    "Please try again in a few seconds."
                )

        st.session_state.chat_messages.append(
            {"role": "assistant", "content": answer}
        )

        st.rerun()