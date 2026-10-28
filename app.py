import re
import io
import html
from datetime import datetime

import fitz
import pandas as pd
import streamlit as st
from docx import Document


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="ResumeAI | Intelligent Resume Screening",
    page_icon="R",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# SKILL DATABASE
# =========================================================

SKILL_DATABASE = {
    "Python": ["python"],
    "Java": ["java"],
    "C": ["c programming", " c language", "c language"],
    "C++": ["c++", "cpp"],
    "JavaScript": ["javascript", "js"],
    "TypeScript": ["typescript"],
    "HTML": ["html"],
    "CSS": ["css"],
    "React": ["react", "react.js", "reactjs"],
    "Angular": ["angular"],
    "Node.js": ["node.js", "nodejs"],
    "PHP": ["php"],
    "Django": ["django"],
    "Flask": ["flask"],
    "SQL": ["sql"],
    "MySQL": ["mysql"],
    "PostgreSQL": ["postgresql"],
    "MongoDB": ["mongodb"],
    "Machine Learning": ["machine learning", "machine-learning"],
    "Deep Learning": ["deep learning", "deep-learning"],
    "Artificial Intelligence": [
        "artificial intelligence",
        "artificial-intelligence"
    ],
    "Generative AI": ["generative ai", "genai", "gen ai"],
    "Natural Language Processing": [
        "natural language processing",
        "nlp"
    ],
    "Large Language Models": [
        "large language models",
        "large language model",
        "llm",
        "llms"
    ],
    "TensorFlow": ["tensorflow"],
    "PyTorch": ["pytorch"],
    "Scikit-learn": ["scikit-learn", "sklearn"],
    "Pandas": ["pandas"],
    "NumPy": ["numpy"],
    "Data Analysis": ["data analysis", "data analytics"],
    "Data Visualization": ["data visualization"],
    "Power BI": ["power bi"],
    "Tableau": ["tableau"],
    "Excel": ["excel", "microsoft excel"],
    "Statistics": ["statistics", "statistical analysis"],
    "Git": ["git"],
    "GitHub": ["github"],
    "Docker": ["docker"],
    "Kubernetes": ["kubernetes"],
    "Jenkins": ["jenkins"],
    "AWS": ["aws", "amazon web services"],
    "Azure": ["microsoft azure", "azure"],
    "Google Cloud": ["google cloud", "gcp"],
    "Linux": ["linux"],
    "REST API": ["rest api", "restful api", "restful apis"],
    "FastAPI": ["fastapi"],
    "API Development": ["api development"],
    "OOP": ["object-oriented programming", "oop"],
    "Data Structures": ["data structures"],
    "Algorithms": ["algorithms", "algorithm design"],
    "Problem Solving": ["problem solving", "problem-solving"],
    "Communication": ["communication skills", "communication"],
    "Teamwork": ["teamwork", "team player"],
    "Agile": ["agile methodology", "agile"],
    "Scrum": ["scrum"],
    "Testing": ["software testing", "testing"],
    "Cybersecurity": ["cybersecurity", "cyber security"],
    "HTML5": ["html5"],
    "CSS3": ["css3"],
}


# =========================================================
# CUSTOM STYLING
# =========================================================

st.markdown(
    """
    <style>
    @import url(
        'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap'
    );

    /* Remove Streamlit's dark top toolbar/header strip. */
    header[data-testid="stHeader"], [data-testid="stToolbar"] {
        display: none !important;
        height: 0 !important;
    }
    #MainMenu, footer { visibility: hidden; }
    .stApp, [data-testid="stAppViewContainer"] {
        background: #f7f8fc !important;
        color: #172033 !important;
        font-family: 'DM Sans', sans-serif;
    }
    [data-testid="stMainBlockContainer"] { padding-top: 1.5rem; }

    h1, h2, h3 {
        font-family: 'Manrope', sans-serif !important;
        color: #172033;
        letter-spacing: -0.5px;
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    [data-testid="stSidebar"] {
        background: #15162e;
        border-right: 1px solid #252642;
    }

    [data-testid="stSidebar"] * {
        color: #f2f2ff;
    }

    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        color: #c5c6e2;
    }

    .brand {
        font-family: 'Manrope', sans-serif;
        font-size: 27px;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: -1px;
        margin-bottom: 3px;
    }

    .brand span {
        color: #a89bff;
    }

    .sidebar-caption {
        font-size: 12px;
        color: #a5a6c6;
        margin-bottom: 28px;
    }

    .top-label {
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 2px;
        color: #7566e8;
        margin-bottom: 8px;
    }

    .hero-title {
        font-family: 'Manrope', sans-serif;
        font-size: clamp(28px, 3vw, 40px);
        line-height: 1.2;
        font-weight: 800;
        color: #191a35;
        margin-bottom: 10px;
    }

    .hero-subtitle {
        color: #777b91;
        font-size: 15px;
        line-height: 1.7;
        margin-bottom: 25px;
    }

    .hero-banner {
        background: linear-gradient(
            120deg,
            #ffffff 0%,
            #f0edff 100%
        );
        border: 1px solid #e7e2ff;
        border-radius: 22px;
        padding: 30px;
        margin-bottom: 26px;
        position: relative;
        overflow: hidden;
    }

    .hero-pill {
        display: inline-block;
        background: #e9e5ff;
        color: #5d4bd6;
        border-radius: 30px;
        padding: 7px 12px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 1px;
        margin-bottom: 15px;
    }

    .metric-card {
        background: #ffffff;
        border: 1px solid #ececf4;
        border-radius: 18px;
        padding: 22px;
        min-height: 132px;
        box-shadow: 0 5px 20px rgba(32, 30, 75, 0.025);
    }

    .metric-label {
        font-size: 13px;
        color: #777b91;
        font-weight: 600;
        margin-bottom: 12px;
    }

    .metric-value {
        font-family: 'Manrope', sans-serif;
        font-size: 32px;
        font-weight: 800;
        color: #191a35;
    }

    .metric-note {
        font-size: 11px;
        color: #85889b;
        margin-top: 6px;
    }

    .section-heading {
        font-family: 'Manrope', sans-serif;
        font-size: 20px;
        font-weight: 800;
        color: #191a35;
        margin-top: 22px;
        margin-bottom: 5px;
    }

    .section-description {
        color: #777b91;
        font-size: 13px;
        margin-bottom: 18px;
    }

    .panel {
        background: #ffffff;
        border: 1px solid #ececf4;
        border-radius: 18px;
        padding: 23px;
        margin-bottom: 18px;
    }

    .candidate-card {
        background: #ffffff;
        border: 1px solid #e9e9f2;
        border-radius: 17px;
        padding: 20px;
        margin-bottom: 12px;
    }

    .candidate-name {
        font-size: 16px;
        font-weight: 800;
        color: #22233f;
        overflow-wrap: anywhere;
    }

    .candidate-meta {
        color: #777b91;
        font-size: 12px;
        margin-top: 5px;
        overflow-wrap: anywhere;
    }

    .score {
        font-family: 'Manrope', sans-serif;
        font-size: 25px;
        font-weight: 800;
        color: #6554df;
    }

    .skill-chip {
        display: inline-block;
        background: #eeebff;
        color: #5948cc;
        padding: 5px 9px;
        border-radius: 7px;
        font-size: 11px;
        font-weight: 700;
        margin: 4px 3px 2px 0;
    }

    .missing-chip {
        display: inline-block;
        background: #fff1e8;
        color: #a95124;
        padding: 5px 9px;
        border-radius: 7px;
        font-size: 11px;
        font-weight: 700;
        margin: 4px 3px 2px 0;
    }

    .empty-state {
        text-align: center;
        padding: 45px 20px;
        background: #ffffff;
        border: 1px dashed #d9d8e9;
        border-radius: 18px;
        color: #777b91;
    }

    .empty-state h3 {
        color: #252642;
        font-size: 20px;
    }

    div.stButton > button {
        border-radius: 10px;
        font-weight: 700;
        min-height: 43px;
        border: 1px solid #e1def8;
    }

    div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #6554df, #8068ef);
        color: white;
        border: none;
    }

    div.stButton > button[kind="primary"]:hover {
        background: linear-gradient(90deg, #5544cf, #7159df);
        color: white;
    }

    div[data-testid="stFileUploader"] {
        background: #fafaff;
        border: 1px dashed #c9c3f5;
        border-radius: 13px;
        padding: 10px;
    }

    /* High-contrast form controls: readable while typing in light mode. */
    div[data-testid="stTextArea"] textarea,
    div[data-testid="stTextInput"] input,
    div[data-testid="stSelectbox"] div[role="combobox"],
    div[data-testid="stNumberInput"] input {
        background: #ffffff !important;
        color: #172033 !important;
        -webkit-text-fill-color: #172033 !important;
        border: 1px solid #d9ddeb !important;
        border-radius: 10px !important;
        caret-color: #6554df !important;
    }
    div[data-testid="stTextArea"] textarea::placeholder,
    div[data-testid="stTextInput"] input::placeholder {
        color: #7b8194 !important;
        -webkit-text-fill-color: #7b8194 !important;
        opacity: 1 !important;
    }
    [data-testid="stWidgetLabel"] label,
    [data-testid="stWidgetLabel"] p,
    .stTextInput label, .stTextArea label,
    .stFileUploader label, .stSelectbox label {
        color: #303650 !important;
        font-weight: 700 !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        background: #ffffff !important;
        color: #303650 !important;
    }
    [data-testid="stFileUploaderDropzone"] * { color: #303650 !important; }
    [data-testid="stDataFrame"] { border: 1px solid #e4e7f0; border-radius: 12px; }
    @media (max-width: 768px) {
        .block-container { padding-left: 1rem; padding-right: 1rem; }
        .metric-value { font-size: 25px; }
    }

    hr {
        border-color: #e9e9f2;
    }

    .footer {
        text-align: center;
        font-size: 11px;
        color: #9699aa;
        padding-top: 30px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# RESUME EXTRACTION FUNCTIONS
# =========================================================

def extract_pdf_text(file_bytes):
    """Extract text from a PDF resume."""
    text_parts = []

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        for page in pdf:
            text_parts.append(page.get_text())

    return "\n".join(text_parts)


def extract_docx_text(file_bytes):
    """Extract paragraphs and tables from a DOCX resume."""
    document = Document(io.BytesIO(file_bytes))

    parts = [paragraph.text for paragraph in document.paragraphs]

    for table in document.tables:
        for row in table.rows:
            parts.append(
                " ".join(cell.text for cell in row.cells)
            )

    return "\n".join(parts)


def extract_resume_text(uploaded_file):
    """Read supported resume formats."""
    file_bytes = uploaded_file.getvalue()
    filename = uploaded_file.name.lower()

    if filename.endswith(".pdf"):
        return extract_pdf_text(file_bytes)

    if filename.endswith(".docx"):
        return extract_docx_text(file_bytes)

    raise ValueError("Unsupported file type. Upload PDF or DOCX.")


# =========================================================
# TEXT PROCESSING FUNCTIONS
# =========================================================

def normalize_text(text):
    """Normalize text for matching."""
    text = text.lower()
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text


def contains_skill(text, skill, aliases):
    """
    Check whether a skill or one of its aliases appears
    as a meaningful phrase in the text.
    """
    normalized = normalize_text(text)

    terms = [skill.lower()] + [
        alias.lower() for alias in aliases
    ]

    for term in terms:
        term = term.strip()

        if not term:
            continue

        # Single-letter programming languages need special handling.
        if term in {"c", "c programming", "c language"}:
            if re.search(r"\bc\s*(?:programming|language)\b", normalized):
                return True
            continue

        if term == "js":
            continue

        pattern = (
            r"(?<![a-z0-9+#])"
            + re.escape(term)
            + r"(?![a-z0-9+#])"
        )

        if re.search(pattern, normalized):
            return True

    return False


def extract_skills(text):
    """Extract skills recognized by the local skill dictionary."""
    found = []

    for skill, aliases in SKILL_DATABASE.items():
        if contains_skill(text, skill, aliases):
            found.append(skill)

    # C and HTML5/CSS3 may otherwise cause duplicate labels.
    return sorted(set(found))


def detect_required_skills(job_description):
    """Detect skills from the job description."""
    return extract_skills(job_description)


def get_email(text):
    """Find the first email address in a resume."""
    match = re.search(
        r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
        text
    )
    return match.group(0) if match else "Not detected"


def get_phone(text):
    """Find a likely phone number."""
    matches = re.findall(
        r"(?<!\d)(?:\+?\d{1,3}[-.\s]?)?"
        r"(?:\(?\d{3,5}\)?[-.\s]?)?"
        r"\d{5}[-.\s]?\d{4}(?!\d)",
        text
    )

    for match in matches:
        digits = re.sub(r"\D", "", match)

        if 10 <= len(digits) <= 13:
            return match.strip()

    return "Not detected"


def get_candidate_name(filename):
    """Use the filename as a dependable initial candidate label."""
    name = filename.rsplit(".", 1)[0]
    name = re.sub(r"[_-]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()

    return name or "Unnamed Candidate"


def extract_education(text):
    """Find likely education phrases using simple keyword rules."""
    education_keywords = [
        "b.tech", "btech", "b.e.", "bachelor of technology",
        "bachelor of engineering", "b.sc", "bsc",
        "bachelor of science", "bca", "mca",
        "m.tech", "mtech", "m.e.", "master of technology",
        "master of science", "m.sc", "msc",
        "mba", "ph.d", "phd", "diploma"
    ]

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    matches = []

    for line in lines:
        lower_line = line.lower()

        if any(keyword in lower_line for keyword in education_keywords):
            if line not in matches:
                matches.append(line)

    return matches[:5]


def calculate_match(required_skills, candidate_skills):
    """
    Calculate a transparent skills-overlap percentage.
    This is not an overall candidate-quality score.
    """
    if not required_skills:
        return 0, [], []

    candidate_set = set(candidate_skills)

    matched = [
        skill for skill in required_skills
        if skill in candidate_set
    ]

    missing = [
        skill for skill in required_skills
        if skill not in candidate_set
    ]

    score = round(
        len(matched) / len(required_skills) * 100
    )

    return score, matched, missing


# =========================================================
# SESSION STATE
# =========================================================

if "candidates" not in st.session_state:
    st.session_state.candidates = []

if "analysis_complete" not in st.session_state:
    st.session_state.analysis_complete = False

if "job_title" not in st.session_state:
    st.session_state.job_title = ""


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:
    st.markdown(
        '<div class="brand">Resume<span>AI</span></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sidebar-caption">'
        'INTELLIGENT TALENT SCREENING'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown("---")

    # Keep the navigation widget's key separate and update it through
    # callbacks. Streamlit does not allow changing a widget's own state
    # after that widget has been instantiated in the current run.
    if "workspace_navigation" not in st.session_state:
        st.session_state.workspace_navigation = "Dashboard"

    page = st.radio(
        "WORKSPACE",
        ["Dashboard", "Screen Resumes", "Candidates", "Reports"],
        key="workspace_navigation",
        label_visibility="visible"
    )

    st.markdown("---")

    st.markdown("**SCREENING WORKSPACE**")
    st.caption("PDF and DOCX resume support")
    st.caption("Skill-based candidate comparison")
    st.caption("CSV report export")

    st.markdown("---")

    st.markdown("**Privacy first**")
    st.caption(
        "Resumes are processed in this running app. "
        "This starter version does not send resume text "
        "to an external AI API."
    )

    st.markdown(
        '<div style="font-size:10px;color:#a5a6c6;'
        'margin-top:35px;">RESUMEAI · PROJECT EDITION</div>',
        unsafe_allow_html=True
    )


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="top-label">RECRUITMENT INTELLIGENCE</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="hero-title">AI Resume Screening</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="hero-subtitle">'
    'Analyze resumes, discover relevant skills, and compare '
    'candidates against your job requirements.'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# METRIC CARDS
# =========================================================

candidates = st.session_state.candidates

total_candidates = len(candidates)

if candidates:
    average_score = round(
        sum(c["match_score"] for c in candidates)
        / total_candidates
    )
else:
    average_score = 0

if candidates:
    highest_score = max(
        c["match_score"] for c in candidates
    )
else:
    highest_score = 0

m1, m2, m3, m4 = st.columns(4)

with m1:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">Resumes analyzed</div>
            <div class="metric-value">{total_candidates}</div>
            <div class="metric-note">In this session</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with m2:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">Average skill match</div>
            <div class="metric-value">{average_score}%</div>
            <div class="metric-note">Across analyzed resumes</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with m3:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">Highest match</div>
            <div class="metric-value">{highest_score}%</div>
            <div class="metric-note">Top skills overlap</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with m4:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">Supported formats</div>
            <div class="metric-value">02</div>
            <div class="metric-note">PDF and DOCX</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# SCREEN RESUMES PAGE
# =========================================================

if page == "Screen Resumes":

    st.markdown(
        '<div class="section-heading">Create a screening session</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-description">'
        'Describe the role, upload resumes, and analyze skill matches.'
        '</div>',
        unsafe_allow_html=True
    )

    with st.form("resume_screening_form"):

        job_title = st.text_input(
            "Job title",
            placeholder="e.g. Python Developer Intern"
        )

        manual_skills = st.text_input(
            "Required skills (comma-separated)",
            placeholder="Python, SQL, Django, Git, REST API",
            help=(
                "Enter the key skills for this role, separated by commas. "
                "If provided, these skills are used for matching. "
                "Leave blank to detect supported skills from the job description."
            )
        )

        job_description = st.text_area(
            "Job description",
            height=180,
            placeholder=(
                "Describe the role, responsibilities, and requirements.\n\n"
                "Example: We are looking for a Python developer with experience "
                "in SQL, Django, REST API development, Git, problem-solving, "
                "and communication."
            ),
            help="Enter the full job description. Text is shown in a readable light input field."
        )

        uploaded_resumes = st.file_uploader(
            "Upload candidate resumes",
            type=["pdf", "docx"],
            accept_multiple_files=True,
            help="You can upload multiple PDF or DOCX resumes."
        )

        analyze_button = st.form_submit_button(
            "Analyze Resumes",
            type="primary",
            use_container_width=True
        )

    if analyze_button:

        if not job_title.strip():
            st.error("Please enter a job title.")

        elif not job_description.strip() and not manual_skills.strip():
            st.error(
                "Enter a job description or provide required skills."
            )

        elif not uploaded_resumes:
            st.error("Please upload at least one resume.")

        else:

            # Required skills can be entered manually or detected.
            if manual_skills.strip():
                requested = [
                    item.strip()
                    for item in manual_skills.split(",")
                    if item.strip()
                ]

                # Map known aliases to canonical skill names.
                required_skills = []

                for requested_skill in requested:
                    normalized_requested = normalize_text(
                        requested_skill
                    )

                    canonical = None

                    for skill, aliases in SKILL_DATABASE.items():
                        possible_names = [
                            skill.lower()
                        ] + [a.lower() for a in aliases]

                        if normalized_requested in possible_names:
                            canonical = skill
                            break

                    final_skill = canonical or requested_skill

                    if final_skill not in required_skills:
                        required_skills.append(final_skill)

            else:
                required_skills = detect_required_skills(
                    job_description
                )

            if not required_skills:
                st.error(
                    "No supported skills were detected. "
                    "Enter required skills in the optional skills field, "
                    "separated by commas."
                )

            else:
                analyzed_candidates = []
                failed_files = []

                progress = st.progress(0)
                status = st.empty()

                for index, uploaded_file in enumerate(uploaded_resumes):

                    status.write(
                        f"Processing {uploaded_file.name}..."
                    )

                    try:
                        resume_text = extract_resume_text(
                            uploaded_file
                        )

                        if not resume_text.strip():
                            failed_files.append(
                                f"{uploaded_file.name}: "
                                "No readable text found. "
                                "The PDF may be scanned or image-based."
                            )
                            progress.progress(
                                (index + 1) / len(uploaded_resumes)
                            )
                            continue

                        candidate_skills = extract_skills(
                            resume_text
                        )

                        score, matched, missing = calculate_match(
                            required_skills,
                            candidate_skills
                        )

                        candidate = {
                            "candidate_name": get_candidate_name(
                                uploaded_file.name
                            ),
                            "filename": uploaded_file.name,
                            "email": get_email(resume_text),
                            "phone": get_phone(resume_text),
                            "skills": candidate_skills,
                            "matched_skills": matched,
                            "missing_skills": missing,
                            "match_score": score,
                            "education": extract_education(
                                resume_text
                            ),
                            "resume_text": resume_text,
                            "required_skills": required_skills,
                            "job_title": job_title.strip(),
                            "analyzed_at": datetime.now().strftime(
                                "%Y-%m-%d %H:%M"
                            )
                        }

                        analyzed_candidates.append(candidate)

                    except Exception as error:
                        failed_files.append(
                            f"{uploaded_file.name}: {str(error)}"
                        )

                    progress.progress(
                        (index + 1) / len(uploaded_resumes)
                    )

                status.empty()
                progress.empty()

                if analyzed_candidates:
                    analyzed_candidates.sort(
                        key=lambda item: item["match_score"],
                        reverse=True
                    )

                    st.session_state.candidates = analyzed_candidates
                    st.session_state.analysis_complete = True
                    st.session_state.job_title = job_title.strip()

                    st.success(
                        f"Successfully analyzed "
                        f"{len(analyzed_candidates)} resume(s)."
                    )

                    st.info(
                        "Open the Candidates page to review the ranking."
                    )

                if failed_files:
                    with st.expander(
                        f"Files that could not be analyzed "
                        f"({len(failed_files)})"
                    ):
                        for message in failed_files:
                            st.warning(message)


# =========================================================
# CANDIDATE RANKING COMPONENT
# =========================================================

def show_candidate_cards(candidate_list):
    """Render candidate results using native Streamlit widgets (not raw HTML)."""
    if not candidate_list:
        st.info("No candidates to display yet. Upload resumes and run a screening session first.")
        return

    for rank, candidate in enumerate(candidate_list, start=1):
        with st.container(border=True):
            left, right = st.columns([4, 1])
            with left:
                st.caption(f"RANK #{rank} · {candidate.get('filename', 'Resume')}")
                st.subheader(candidate.get("candidate_name", "Unnamed Candidate"))
                st.caption(candidate.get("email", "Email not detected"))
                st.markdown("**Matched required skills**")
                matched = candidate.get("matched_skills", [])
                st.write(", ".join(matched) if matched else "No required skills detected in this resume.")
                st.markdown("**Required skills not detected**")
                missing = candidate.get("missing_skills", [])
                st.write(", ".join(missing) if missing else "All listed required skills were detected.")
            with right:
                st.metric("Skills match", f"{candidate.get('match_score', 0)}%")
                st.progress(max(0, min(100, int(candidate.get("match_score", 0)))) / 100)


# =========================================================
# DASHBOARD PAGE
# =========================================================

if page == "Dashboard":

    st.markdown(
        """
        <div class="hero-banner">
            <div class="hero-pill">SMARTER HIRING STARTS HERE</div>
            <div style="font-family:'Manrope',sans-serif;
            font-size:28px;font-weight:800;color:#191a35;">
                Find the right skills.<br>
                Discover relevant candidates.
            </div>
            <p style="color:#777b91;font-size:14px;line-height:1.7;
            max-width:650px;margin-top:12px;">
                Screen multiple resumes against a job description,
                identify relevant skills, and compare candidates
                through a transparent skills-match score.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    left, right = st.columns([1.2, 0.8])

    with left:
        st.markdown(
            '<div class="section-heading">Your screening workspace</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="section-description">'
            'Start a new screening session or review existing results.'
            '</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <div class="panel">
                <div style="font-size:16px;font-weight:800;color:#22233f;">
                    01 &nbsp; Upload resumes
                </div>
                <p style="font-size:13px;color:#777b91;">
                    Import multiple PDF or DOCX resumes into one session.
                </p>
                <div style="font-size:16px;font-weight:800;color:#22233f;
                margin-top:18px;">
                    02 &nbsp; Extract candidate skills
                </div>
                <p style="font-size:13px;color:#777b91;">
                    Identify recognized technical and professional skills.
                </p>
                <div style="font-size:16px;font-weight:800;color:#22233f;
                margin-top:18px;">
                    03 &nbsp; Compare and rank
                </div>
                <p style="font-size:13px;color:#777b91;">
                    Compare detected skills with the job requirements.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        def open_screening_page():
            # Streamlit executes on_click callbacks before rebuilding widgets.
            st.session_state["workspace_navigation"] = "Screen Resumes"

        st.button(
            "Start Resume Screening",
            type="primary",
            use_container_width=True,
            on_click=open_screening_page
        )

    with right:
        st.markdown(
            '<div class="section-heading">Screening insights</div>',
            unsafe_allow_html=True
        )

        if candidates:
            st.markdown(
                '<div class="panel">'
                '<div style="font-size:14px;font-weight:800;color:#22233f;">'
                'Latest screening session</div>'
                f'<p style="color:#777b91;font-size:13px;">'
                f'Role: {html.escape(st.session_state.job_title)}</p>'
                f'<p style="color:#6554df;font-size:27px;font-weight:800;">'
                f'{len(candidates)} candidates</p>'
                '<p style="color:#777b91;font-size:13px;">'
                'Results are available in this running session.</p>'
                '</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                """
                <div class="panel">
                    <div style="font-size:15px;font-weight:800;
                    color:#22233f;">
                        No screening session yet
                    </div>
                    <p style="color:#777b91;font-size:13px;
                    line-height:1.7;">
                        Upload resumes and a job description to populate
                        your dashboard with real screening results.
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.markdown(
            """
            <div class="panel">
                <div style="font-size:14px;font-weight:800;color:#22233f;">
                    How scoring works
                </div>
                <p style="color:#777b91;font-size:13px;line-height:1.8;">
                    Match score = required skills detected in the resume
                    divided by total required skills, multiplied by 100.
                </p>
                <p style="color:#a95124;font-size:12px;">
                    This is a preliminary skills-overlap indicator,
                    not a measure of a candidate's overall ability.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )


# =========================================================
# CANDIDATES PAGE
# =========================================================

if page == "Candidates":

    st.markdown(
        '<div class="section-heading">Candidate ranking</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-description">'
        'Review detected skills, compare match scores, and inspect '
        'individual candidate details.'
        '</div>',
        unsafe_allow_html=True
    )

    if candidates:

        search_query = st.text_input(
            "Search candidates",
            placeholder="Search by candidate or filename"
        )

        filtered_candidates = [
            candidate for candidate in candidates
            if search_query.lower() in
            (
                candidate["candidate_name"] + " " +
                candidate["filename"] + " " +
                candidate["email"]
            ).lower()
        ]

        show_candidate_cards(filtered_candidates)

        st.markdown(
            '<div class="section-heading">Candidate details</div>',
            unsafe_allow_html=True
        )

        candidate_options = {
            f"{candidate['candidate_name']} — {candidate['filename']} (#{index + 1})": candidate
            for index, candidate in enumerate(filtered_candidates)
        }
        selected_label = st.selectbox(
            "Select a candidate",
            options=list(candidate_options.keys()),
            key="candidate_details_select"
        ) if candidate_options else None

        if selected_label:
            selected = candidate_options[selected_label]

            st.markdown(
                f"""
                <div class="panel">
                    <div style="font-size:22px;font-weight:800;
                    color:#22233f;">
                        {html.escape(selected["candidate_name"])}
                    </div>
                    <p style="font-size:13px;color:#777b91;">
                        {html.escape(selected["filename"])}
                    </p>
                    <hr>
                    <b>Email:</b> {html.escape(selected["email"])}<br>
                    <b>Phone:</b> {html.escape(selected["phone"])}<br>
                    <b>Skills-match score:</b>
                    {selected["match_score"]}%
                </div>
                """,
                unsafe_allow_html=True
            )

            d1, d2 = st.columns(2)

            with d1:
                st.markdown("**Detected skills**")

                if selected["skills"]:
                    st.write(", ".join(selected["skills"]))
                else:
                    st.write("No skills from the current dictionary detected.")

                st.markdown("**Education phrases detected**")

                if selected["education"]:
                    for education_item in selected["education"]:
                        st.write(f"- {education_item}")
                else:
                    st.write("No recognizable education phrase detected.")

            with d2:
                st.markdown("**Matched required skills**")
                st.write(
                    ", ".join(selected["matched_skills"]) or "None detected"
                )

                st.markdown("**Required skills not detected**")
                st.write(
                    ", ".join(selected["missing_skills"]) or "None"
                )

            with st.expander("View extracted resume text"):
                st.text(selected["resume_text"])

            st.caption(
                "Extraction is heuristic. Missing information may still "
                "be present in the resume but not detected by the program."
            )

    else:
        show_candidate_cards([])


# =========================================================
# REPORTS PAGE
# =========================================================

if page == "Reports":

    st.markdown(
        '<div class="section-heading">Screening reports</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-description">'
        'Review the screening summary and export results.'
        '</div>',
        unsafe_allow_html=True
    )

    if candidates:

        report_rows = []

        for rank, candidate in enumerate(candidates, start=1):
            report_rows.append({
                "Rank": rank,
                "Candidate": candidate["candidate_name"],
                "Filename": candidate["filename"],
                "Email": candidate["email"],
                "Phone": candidate["phone"],
                "Match Score (%)": candidate["match_score"],
                "Matched Skills": ", ".join(
                    candidate["matched_skills"]
                ),
                "Required Skills Not Detected": ", ".join(
                    candidate["missing_skills"]
                ),
                "Detected Skills": ", ".join(
                    candidate["skills"]
                ),
                "Job Title": candidate["job_title"],
                "Analyzed At": candidate["analyzed_at"]
            })

        report_df = pd.DataFrame(report_rows)

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
        )

        csv_data = report_df.to_csv(index=False).encode("utf-8-sig")

        st.download_button(
            label="Download Screening Report (CSV)",
            data=csv_data,
            file_name="resume_screening_report.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True
        )

        st.info(
            "The report contains extracted information and skills-match "
            "scores. Review each resume before making hiring decisions."
        )

    else:
        show_candidate_cards([])


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer">
        RESUMEAI · INTELLIGENT RESUME SCREENING SYSTEM<br>
        Designed to support human review, not replace hiring decisions.
    </div>
    """,
    unsafe_allow_html=True
)