import os
import re
import json
from pathlib import Path
from collections import Counter

import streamlit as st
import pandas as pd

from pypdf import PdfReader
from docx import Document

# Groq is optional until the API key is configured
try:
    from groq import Groq
except ImportError:
    Groq = None


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="SCOCOEX NEXUS",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).parent
DOCS_DIR = ROOT / "docs"
COMPANIES_FILE = ROOT / "companies.json"

SUPPORTED_EXTENSIONS = {".pdf", ".docx"}


# ============================================================
# PREMIUM UI
# ============================================================

st.markdown(
    """
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

[data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at 10% 0%, rgba(88, 80, 255, 0.16), transparent 28%),
        radial-gradient(circle at 90% 10%, rgba(0, 190, 255, 0.10), transparent 25%),
        #070A12;
    color: #F7F8FC;
}

[data-testid="stSidebar"] {
    background: rgba(8, 11, 20, 0.98);
    border-right: 1px solid rgba(255,255,255,0.07);
}

.hero {
    padding: 52px 42px;
    border-radius: 30px;
    border: 1px solid rgba(255,255,255,0.10);
    background:
        linear-gradient(
            135deg,
            rgba(28,32,58,0.96),
            rgba(10,13,25,0.92)
        );
    margin-bottom: 30px;
}

.hero-label {
    color: #9AA8FF;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: 0.18em;
    text-transform: uppercase;
}

.hero h1 {
    font-size: 54px;
    line-height: 1.0;
    margin: 12px 0;
    font-weight: 800;
}

.hero p {
    color: #B2B9CB;
    font-size: 17px;
    max-width: 800px;
    line-height: 1.7;
}

.card {
    padding: 22px;
    border-radius: 20px;
    background: rgba(255,255,255,0.045);
    border: 1px solid rgba(255,255,255,0.08);
    margin-bottom: 14px;
}

.metric-card {
    padding: 24px;
    border-radius: 20px;
    background: rgba(255,255,255,0.045);
    border: 1px solid rgba(255,255,255,0.08);
}

.source {
    padding: 12px 15px;
    border-radius: 12px;
    background: rgba(125,110,255,0.10);
    border: 1px solid rgba(125,110,255,0.18);
    margin-top: 8px;
}

.small {
    color: #929AAF;
    font-size: 13px;
}

.answer {
    padding: 25px;
    border-radius: 20px;
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    line-height: 1.8;
}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# DOCUMENT READING
# ============================================================

def extract_pdf(path: Path):
    """
    Extract text from PDF while preserving page numbers.
    """
    reader = PdfReader(str(path))

    pages = []

    for page_number, page in enumerate(reader.pages, start=1):

        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""

        if text.strip():
            pages.append(
                {
                    "page": page_number,
                    "text": text.strip(),
                }
            )

    return pages


def extract_docx(path: Path):
    """
    Extract paragraphs and tables from DOCX.
    """
    document = Document(str(path))

    sections = []

    for paragraph in document.paragraphs:

        text = paragraph.text.strip()

        if text:
            sections.append(
                {
                    "page": None,
                    "text": text,
                }
            )

    # Tables
    for table_index, table in enumerate(document.tables, start=1):

        rows = []

        for row in table.rows:

            cells = [
                cell.text.strip()
                for cell in row.cells
            ]

            rows.append(" | ".join(cells))

        if rows:
            sections.append(
                {
                    "page": None,
                    "text": f"[TABLE {table_index}]\n" + "\n".join(rows),
                }
            )

    return sections


def read_document(path: Path):

    extension = path.suffix.lower()

    if extension == ".pdf":
        return extract_pdf(path)

    if extension == ".docx":
        return extract_docx(path)

    return []


# ============================================================
# DOCUMENT INDEX
# ============================================================

@st.cache_data(show_spinner=False)
def load_all_documents():

    documents = []

    if not DOCS_DIR.exists():
        return documents

    for path in sorted(DOCS_DIR.rglob("*")):

        if not path.is_file():
            continue

        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue

        try:

            sections = read_document(path)

            documents.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
                    "sections": sections,
                }
            )

        except Exception as error:

            documents.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
                    "sections": [],
                    "error": str(error),
                }
            )

    return documents


# ============================================================
# CHUNKING
# ============================================================

def normalize_text(text):

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def create_chunks(documents, chunk_size=1400):

    chunks = []

    for document in documents:

        name = document["name"]

        for section in document["sections"]:

            text = normalize_text(section["text"])

            if not text:
                continue

            page = section.get("page")

            # Simple character-based chunking
            start = 0

            while start < len(text):

                chunk_text = text[start:start + chunk_size]

                chunks.append(
                    {
                        "document": name,
                        "page": page,
                        "text": chunk_text,
                    }
                )

                start += chunk_size

    return chunks


# ============================================================
# SIMPLE LOCAL RETRIEVAL
# ============================================================

def tokenize(text):

    return re.findall(
        r"[\w\u0600-\u06FF\u4e00-\u9fff\u0400-\u04FF]+",
        text.lower(),
    )


def retrieve_chunks(question, chunks, top_k=6):

    question_tokens = set(tokenize(question))

    if not question_tokens:
        return []

    scored = []

    for chunk in chunks:

        chunk_tokens = tokenize(chunk["text"])

        if not chunk_tokens:
            continue

        counts = Counter(chunk_tokens)

        score = 0

        for token in question_tokens:
            score += counts.get(token, 0)

        if score > 0:
            scored.append(
                (
                    score,
                    chunk,
                )
            )

    scored.sort(
        key=lambda x: x[0],
        reverse=True,
    )

    return [
        item[1]
        for item in scored[:top_k]
    ]


# ============================================================
# GROQ
# ============================================================

def get_groq_client():

    api_key = None

    # Streamlit secrets first
    try:
        api_key = st.secrets.get("GROQ_API_KEY")
    except Exception:
        pass

    # Environment variable
    if not api_key:
        api_key = os.getenv("GROQ_API_KEY")

    if not api_key:
        return None

    if Groq is None:
        return None

    return Groq(api_key=api_key)


def ask_groq(
    question,
    context,
    language="same language as the user",
):

    client = get_groq_client()

    if client is None:
        return None

    system_prompt = f"""
You are SCOCOEX NEXUS AI.

You are an event intelligence and knowledge assistant.

IMPORTANT RULES:

1. Answer primarily from the provided SCOCOEX documents.
2. Do not invent facts that are not supported by the documents.
3. If the documents do not contain enough information, clearly say so.
4. Preserve company names, project names and terminology from the documents.
5. Answer in {language}.
6. Be professional and suitable for investors, companies and international delegates.
7. When possible, mention the source document.
"""

    user_prompt = f"""
QUESTION:

{question}

DOCUMENT CONTEXT:

{context}

Provide a clear and useful answer.
"""

    try:

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": user_prompt,
                },
            ],
            temperature=0.2,
            max_tokens=1800,
        )

        return response.choices[0].message.content

    except Exception as error:

        return f"Groq error: {error}"


# ============================================================
# SUMMARY
# ============================================================

def summarize_document(document):

    sections = document.get("sections", [])

    text = "\n".join(
        section["text"]
        for section in sections
    )

    text = text[:18000]

    if not text:
        return "No readable text was found."

    client = get_groq_client()

    if client is None:
        return (
            "Groq API is not configured yet. "
            "The document was successfully read and is ready for AI summarization."
        )

    prompt = f"""
Summarize the following SCOCOEX document.

Return:

1. Executive Summary
2. Key Points
3. Important Organizations / Companies
4. Important Numbers
5. Opportunities
6. Strategic Takeaways

Document:

{text}
"""

    try:

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "system",
                    "content": "You are a professional strategic intelligence analyst.",
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            temperature=0.2,
            max_tokens=2200,
        )

        return response.choices[0].message.content

    except Exception as error:

        return f"Groq error: {error}"


# ============================================================
# COMPANIES
# ============================================================

@st.cache_data
def load_companies():

    if not COMPANIES_FILE.exists():
        return []

    try:

        with open(
            COMPANIES_FILE,
            "r",
            encoding="utf-8",
        ) as file:

            return json.load(file)

    except Exception:

        return []


# ============================================================
# LANGUAGE
# ============================================================

LANGUAGES = {
    "English": "English",
    "فارسی": "Persian",
    "العربية": "Arabic",
    "中文": "Chinese",
    "Русский": "Russian",
}


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    """
    <div style="padding:10px 0 25px 0;">
        <div style="font-size:25px;font-weight:800;">
            🌐 SCOCOEX
        </div>
        <div style="color:#8D98AF;font-size:13px;">
            NEXUS Intelligence Platform
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

page = st.sidebar.radio(
    "EXPLORE",
    [
        "Home",
        "Information Center",
        "AI Assistant",
        "International Companies",
        "Intelligence",
        "Gallery",
        "About",
        "Contact",
    ],
    label_visibility="collapsed",
)

st.sidebar.divider()

language = st.sidebar.selectbox(
    "AI Response Language",
    list(LANGUAGES.keys()),
)

documents = load_all_documents()
chunks = create_chunks(documents)
companies = load_companies()


# ============================================================
# HOME
# ============================================================

if page == "Home":

    st.markdown(
        """
        <div class="hero">

            <div class="hero-label">
                SCOCOEX NEXUS
            </div>

            <h1>
                Connect.<br>
                Discover.<br>
                Invest.
            </h1>

            <p>
                An AI-powered global intelligence and knowledge
                platform designed to connect organizations,
                investors, companies and strategic information.
            </p>

        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric(
            "Documents",
            len(documents),
        )

    with c2:
        st.metric(
            "Knowledge Chunks",
            len(chunks),
        )

    with c3:
        st.metric(
            "International Companies",
            len(companies),
        )

    with c4:
        st.metric(
            "Languages",
            "5",
        )

    st.markdown("### Intelligence Infrastructure")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            """
            <div class="metric-card">
            <h3>📚 Knowledge</h3>
            <p>
            PDF and Word documents become a searchable
            institutional knowledge base.
            </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:

        st.markdown(
            """
            <div class="metric-card">
            <h3>🤖 AI</h3>
            <p>
            Groq-powered AI answers questions using
            information retrieved from SCOCOEX documents.
            </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col3:

        st.markdown(
            """
            <div class="metric-card">
            <h3>🌍 Network</h3>
            <p>
            Explore international companies and strategic
            organizations connected to the ecosystem.
            </p>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# INFORMATION CENTER
# ============================================================

elif page == "Information Center":

    st.title("📚 Information Center")

    st.write(
        "All PDF and DOCX files located inside the `docs/` folder "
        "are automatically indexed by the application."
    )

    if not DOCS_DIR.exists():

        st.error(
            "The docs folder does not exist. "
            "Create a folder named `docs` in the repository."
        )

    elif not documents:

        st.warning(
            "No PDF or DOCX files were found inside `docs/`."
        )

    else:

        st.success(
            f"{len(documents)} document(s) detected."
        )

        search = st.text_input(
            "🔎 Search documents",
            placeholder="Search by filename...",
        )

        filtered_documents = documents

        if search:

            filtered_documents = [
                doc
                for doc in documents
                if search.lower()
                in doc["name"].lower()
            ]

        for document in filtered_documents:

            with st.expander(
                f"📄 {document['name']}"
            ):

                section_count = len(
                    document.get("sections", [])
                )

                st.write(
                    f"Sections extracted: **{section_count}**"
                )

                if document.get("error"):

                    st.error(
                        document["error"]
                    )

                if st.button(
                    "Generate AI Summary",
                    key="summary_" + document["name"],
                ):

                    with st.spinner(
                        "Analyzing document..."
                    ):

                        summary = summarize_document(
                            document
                        )

                    st.markdown(
                        '<div class="answer">',
                        unsafe_allow_html=True,
                    )

                    st.markdown(summary)

                    st.markdown(
                        "</div>",
                        unsafe_allow_html=True,
                    )


# ============================================================
# AI ASSISTANT
# ============================================================

elif page == "AI Assistant":

    st.title("🤖 SCOCOEX AI Assistant")

    st.caption(
        "Ask questions about the information contained in the SCOCOEX Knowledge Base."
    )

    question = st.text_area(
        "Your question",
        placeholder=(
            "مثلاً: مهم‌ترین فرصت‌های سرمایه‌گذاری "
            "که در اسناد SCOCOEX ذکر شده چیست؟"
        ),
        height=130,
    )

    if st.button(
        "Ask SCOCOEX AI",
        type="primary",
        use_container_width=True,
    ):

        if not question.strip():

            st.warning(
                "Please enter a question."
            )

        elif not chunks:

            st.error(
                "No documents are available for the AI Knowledge Base."
            )

        else:

            with st.spinner(
                "Searching the SCOCOEX Knowledge Base..."
            ):

                results = retrieve_chunks(
                    question,
                    chunks,
                    top_k=6,
                )

            if not results:

                st.warning(
                    "I could not find relevant information "
                    "in the current Knowledge Base."
                )

            else:

                context_parts = []

                for result in results:

                    source = result["document"]

                    page = result.get("page")

                    if page:

                        source_info = (
                            f"{source} — Page {page}"
                        )

                    else:

                        source_info = source

                    context_parts.append(
                        f"""
SOURCE: {source_info}

CONTENT:
{result['text']}
"""
                    )

                context = "\n\n".join(
                    context_parts
                )

                response_language = LANGUAGES[
                    language
                ]

                with st.spinner(
                    "Generating intelligence..."
                ):

                    answer = ask_groq(
                        question,
                        context,
                        response_language,
                    )

                if answer is None:

                    st.warning(
                        "Groq API is not configured yet."
                    )

                    st.info(
                        "The document retrieval system is working. "
                        "Add GROQ_API_KEY in Streamlit Secrets to activate AI answers."
                    )

                else:

                    st.markdown(
                        '<div class="answer">',
                        unsafe_allow_html=True,
                    )

                    st.markdown(answer)

                    st.markdown(
                        "</div>",
                        unsafe_allow_html=True,
                    )

                    st.markdown("### Sources")

                    shown_sources = set()

                    for result in results:

                        key = (
                            result["document"],
                            result.get("page"),
                        )

                        if key in shown_sources:
                            continue

                        shown_sources.add(key)

                        page = result.get("page")

                        if page:

                            text = (
                                f"📄 {result['document']} "
                                f"— Page {page}"
                            )

                        else:

                            text = (
                                f"📄 {result['document']}"
                            )

                        st.markdown(
                            f'<div class="source">{text}</div>',
                            unsafe_allow_html=True,
                        )


# ============================================================
# INTERNATIONAL COMPANIES
# ============================================================

elif page == "International Companies":

    st.title("🌍 International Companies")

    if not companies:

        st.warning(
            "companies.json was not found or is empty."
        )

    else:

        search = st.text_input(
            "🔎 Search company"
        )

        filtered = companies

        if search:

            filtered = [
                company
                for company in companies
                if search.lower()
                in json.dumps(
                    company,
                    ensure_ascii=False
                ).lower()
            ]

        st.caption(
            f"{len(filtered)} organizations"
        )

        columns = st.columns(3)

        for index, company in enumerate(filtered):

            with columns[index % 3]:

                name = company.get(
                    "name",
                    "International Company"
                )

                website = company.get(
                    "website",
                    "#"
                )

                sector = company.get(
                    "sector",
                    "International Organization"
                )

                st.markdown(
                    f"""
                    <div class="card">

                    <h3>{name}</h3>

                    <div class="small">
                    {sector}
                    </div>

                    <br>

                    <a href="{website}" target="_blank">
                    Visit official website →
                    </a>

                    </div>
                    """,
                    unsafe_allow_html=True,
                )


# ============================================================
# INTELLIGENCE
# ============================================================

elif page == "Intelligence":

    st.title("📊 Intelligence Dashboard")

    st.caption(
        "Initial document intelligence layer."
    )

    if not documents:

        st.info(
            "Upload documents to the docs/ folder first."
        )

    else:

        rows = []

        for document in documents:

            full_text = " ".join(
                section["text"]
                for section in document["sections"]
            )

            words = tokenize(full_text)

            numbers = re.findall(
                r"\b\d+(?:[.,]\d+)?\b",
                full_text,
            )

            rows.append(
                {
                    "Document": document["name"],
                    "Words": len(words),
                    "Numeric values": len(numbers),
                    "Sections": len(
                        document["sections"]
                    ),
                }
            )

        df = pd.DataFrame(rows)

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Document Size")

        chart_df = df.set_index(
            "Document"
        )["Words"]

        st.bar_chart(chart_df)

        st.markdown("### Numeric Information")

        numeric_df = df.set_index(
            "Document"
        )["Numeric values"]

        st.bar_chart(numeric_df)

        st.info(
            "The next intelligence layer can extract named entities, "
            "financial figures, countries, sectors and investment amounts "
            "and convert them into structured dashboards."
        )


# ============================================================
# GALLERY
# ============================================================

elif page == "Gallery":

    st.title("🖼️ SCOCOEX Gallery")

    st.info(
        "Event images can be added here later."
    )

    st.write(
        "Recommended future sections:"
    )

    st.markdown(
        """
        - Event Highlights
        - Speakers
        - Meetings
        - International Delegations
        - Partners
        - Investment Sessions
        """
    )


# ============================================================
# ABOUT
# ============================================================

elif page == "About":

    st.title("ℹ️ About SCOCOEX")

    st.markdown(
        """
        ### SCOCOEX NEXUS

        SCOCOEX NEXUS is designed as an AI-powered
        global intelligence and knowledge platform.

        The platform brings together:

        **Knowledge**

        Institutional documents, reports and strategic information.

        **Artificial Intelligence**

        AI-powered search, question answering and document analysis.

        **International Network**

        Companies, organizations, investors and strategic partners.

        **Intelligence**

        Structured information, statistics and analytical insights.
        """
    )


# ============================================================
# CONTACT
# ============================================================

elif page == "Contact":

    st.title("📩 Connect with SCOCOEX")

    with st.form("contact_form"):

        name = st.text_input("Name")

        company = st.text_input(
            "Company / Organization"
        )

        email = st.text_input("Email")

        interest = st.selectbox(
            "Interest",
            [
                "Investment",
                "Partnership",
                "International Company",
                "Government",
                "Technology",
                "Media",
                "Other",
            ],
        )

        message = st.text_area(
            "Message",
            height=160,
        )

        submitted = st.form_submit_button(
            "Send Inquiry",
            type="primary",
            use_container_width=True,
        )

        if submitted:

            st.success(
                "Thank you. Your inquiry has been received."
                )
