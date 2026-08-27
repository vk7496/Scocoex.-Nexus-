import streamlit as st
from pathlib import Path
import json
from urllib.parse import urlparse

st.set_page_config(
    page_title="SCOCOEX | Global Intelligence",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).parent
DATA = ROOT / "data" / "companies.json"

@st.cache_data
def load_companies():
    return json.loads(DATA.read_text(encoding="utf-8"))

companies = load_companies()

# ---------- premium visual system ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] { font-family: Inter, sans-serif; }

[data-testid="stAppViewContainer"] {
    background:
      radial-gradient(circle at 10% 0%, rgba(91, 72, 255, .14), transparent 28%),
      radial-gradient(circle at 90% 10%, rgba(0, 190, 255, .10), transparent 24%),
      #070A12;
    color: #F7F8FC;
}
[data-testid="stSidebar"] {
    background: rgba(9, 12, 22, .96);
    border-right: 1px solid rgba(255,255,255,.08);
}
.hero {
    padding: 42px 36px;
    border: 1px solid rgba(255,255,255,.10);
    border-radius: 28px;
    background: linear-gradient(135deg, rgba(22,26,48,.94), rgba(10,13,25,.78));
    box-shadow: 0 24px 80px rgba(0,0,0,.28);
    margin-bottom: 28px;
}
.eyebrow { color:#8FA4FF; font-weight:700; letter-spacing:.16em; text-transform:uppercase; font-size:12px; }
.hero h1 { font-size:48px; line-height:1.02; margin:10px 0; }
.hero p { color:#AEB7CC; font-size:17px; max-width:760px; }
.metric {
    padding:20px; border-radius:20px; background:rgba(255,255,255,.045);
    border:1px solid rgba(255,255,255,.08);
}
.metric .n { font-size:32px; font-weight:800; }
.metric .l { color:#9CA6BB; font-size:13px; }
.company-card {
    min-height:245px; padding:22px; border-radius:22px;
    background:linear-gradient(180deg, rgba(255,255,255,.065), rgba(255,255,255,.025));
    border:1px solid rgba(255,255,255,.09);
    box-shadow:0 12px 38px rgba(0,0,0,.18);
}
.company-logo {
    width:58px; height:58px; border-radius:16px; background:#fff;
    object-fit:contain; padding:8px; margin-bottom:14px;
}
.company-name { font-size:20px; font-weight:750; margin-bottom:8px; }
.badge {
    display:inline-block; padding:5px 9px; border-radius:999px;
    background:rgba(143,164,255,.12); color:#B9C4FF; font-size:11px; margin:2px;
}
a { color:#AFC0FF !important; text-decoration:none; }
</style>
""", unsafe_allow_html=True)

# ---------- custom navigation ----------
st.sidebar.markdown("## 🌐 SCOCOEX")
st.sidebar.caption("Global Intelligence Portal")
page = st.sidebar.radio(
    "EXPLORE",
    ["Home", "International Companies", "Information Center", "AI Assistant", "Intelligence", "Gallery", "About", "Contact"],
    label_visibility="collapsed",
)

if page == "Home":
    st.markdown("""
    <div class="hero">
      <div class="eyebrow">Global Intelligence Platform</div>
      <h1>Connect. Discover.<br>Invest.</h1>
      <p>SCOCOEX transforms international knowledge, organizations and strategic information into an intelligent digital experience.</p>
    </div>
    """, unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    for col, n, label in [
        (c1, len(companies), "International organizations"),
        (c2, "AI", "Knowledge assistant"),
        (c3, "24/7", "Information access"),
        (c4, "∞", "Growth potential"),
    ]:
        with col:
            st.markdown(f'<div class="metric"><div class="n">{n}</div><div class="l">{label}</div></div>', unsafe_allow_html=True)
    st.markdown("### Intelligence at the center")
    st.write("Explore the international network, institutional documents, AI-powered Q&A and event intelligence from one portal.")

elif page == "International Companies":
    st.markdown("## 🌍 International Network")
    st.caption("Organizations and companies supplied for the SCOCOEX international network.")
    q = st.text_input("🔎 Search", placeholder="Search company or website...")
    filtered = [c for c in companies if not q or q.lower() in (c["name"]+" "+c["website"]).lower()]
    st.caption(f"{len(filtered)} organizations")
    cols = st.columns(3)
    for i, c in enumerate(filtered):
        with cols[i % 3]:
            logo = c.get("logo_url","")
            name = c.get("name","Company")
            website = c.get("website","#")
            st.markdown(f"""
            <div class="company-card">
              <img class="company-logo" src="{logo}" onerror="this.style.display='none';">
              <div class="company-name">{name}</div>
              <span class="badge">International</span>
              {f'<span class="badge">{c["sector"]}</span>' if c.get("sector") else ''}
              <p style="color:#9CA6BB;font-size:13px;margin-top:14px;">{c.get("description","")}</p>
              <a href="{website}" target="_blank">Visit official website →</a>
            </div>
            """, unsafe_allow_html=True)

elif page == "Information Center":
    st.markdown("## 📚 Information Center")
    st.info("Put PDF and DOCX files inside `knowledge_base/documents/`. The next module will index them for search, summaries, statistics and RAG Q&A.")
    docs = list((ROOT/"knowledge_base"/"documents").glob("*"))
    if docs:
        for d in docs:
            st.write(f"📄 **{d.name}**")
    else:
        st.warning("No documents found yet.")

elif page == "AI Assistant":
    st.markdown("## 🤖 SCOCOEX AI Assistant")
    st.info("The RAG layer will answer from the documents placed in `knowledge_base/documents/`, with source citations.")
    st.text_area("Ask a question", placeholder="What are the main investment opportunities?", height=130)
    st.button("Ask SCOCOEX AI", type="primary", use_container_width=True)

elif page == "Intelligence":
    st.markdown("## 📊 Intelligence")
    st.info("This dashboard will visualize statistics extracted from the Knowledge Base.")
    st.metric("Companies in network", len(companies))

elif page == "Gallery":
    st.markdown("## 🖼️ Gallery")
    st.info("Place event images in `assets/images/` when ready.")

elif page == "About":
    st.markdown("## About SCOCOEX")
    st.write("This section is ready for the official SCOCOEX profile, mission, vision and leadership information.")

elif page == "Contact":
    st.markdown("## 📩 Connect with SCOCOEX")
    with st.form("contact"):
        name = st.text_input("Name")
        company = st.text_input("Company")
        email = st.text_input("Email")
        interest = st.selectbox("Interest", ["Investor", "International Company", "Government", "Partner", "Media", "General"])
        message = st.text_area("Message", height=150)
        st.form_submit_button("Send Inquiry", type="primary")
