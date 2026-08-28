
import os
import re
import json
import html
from pathlib import Path
from collections import Counter

import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document

try:
    from groq import Groq
except ImportError:
    Groq = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
except ImportError:
    SimpleDocTemplate = None


# ============================================================
# APP CONFIG
# ============================================================

st.set_page_config(
    page_title="SCOCOEX NEXUS",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent
DOCS_DIR = ROOT / "docs"
COMPANIES_FILE = ROOT / "companies.json"

# Logo: accepts assets/scocoex_logo.png, assets/scocoex_logo-1.png,
# or any other PNG beginning with scocoex_logo.
LOGO_FILES = list((ROOT / "assets").glob("scocoex_logo*.png")) if (ROOT / "assets").exists() else []
LOGO_PATH = LOGO_FILES[0] if LOGO_FILES else None

SUPPORTED_EXTENSIONS = {".pdf", ".docx"}

LANGUAGES = {
    "English": "en",
    "فارسی": "fa",
    "العربية": "ar",
    "中文": "zh",
    "Русский": "ru",
}

# ============================================================
# TRANSLATIONS
# ============================================================

T = {
    "en": {
        "nav": "NAVIGATION",
        "home": "Home",
        "info": "Information Center",
        "assistant": "AI Assistant",
        "companies": "International Companies",
        "intel": "Intelligence",
        "gallery": "Gallery",
        "about": "About",
        "contact": "Contact",
        "language": "Language",
        "response_language": "AI Response Language",
        "hero_label": "GLOBAL INVESTMENT & INTELLIGENCE PLATFORM",
        "hero_title": "Connect.<br>Discover.<br>Invest.",
        "hero_text": "An AI-powered global intelligence and knowledge platform designed to connect organizations, investors, companies and strategic information.",
        "documents": "Documents",
        "chunks": "Knowledge Chunks",
        "international_companies": "International Companies",
        "languages": "Languages",
        "knowledge": "Knowledge",
        "knowledge_text": "Institutional documents, reports and strategic information become a searchable knowledge base.",
        "ai": "Artificial Intelligence",
        "ai_text": "Groq-powered AI answers questions using information retrieved from SCOCOEX documents.",
        "network": "Global Network",
        "network_text": "Explore international companies, investors and strategic organizations.",
        "information_title": "Information Center",
        "information_text": "PDF and DOCX files inside the docs/ folder are automatically indexed.",
        "search_documents": "Search documents",
        "sections": "Sections extracted",
        "summary": "Generate AI Summary",
        "summary_title": "AI Document Summary",
        "ask_title": "SCOCOEX AI Assistant",
        "ask_caption": "Ask questions about the information contained in the SCOCOEX Knowledge Base.",
        "your_question": "Your question",
        "ask_button": "Ask SCOCOEX AI",
        "searching": "Searching the SCOCOEX Knowledge Base...",
        "generating": "Generating intelligence...",
        "sources": "Sources",
        "companies_title": "International Companies",
        "search_company": "Search company",
        "visit": "Visit official website →",
        "company_count": "organizations",
        "intel_title": "Intelligence Dashboard",
        "intel_caption": "Initial document intelligence layer.",
        "document_size": "Document Size",
        "numeric": "Numeric Information",
        "about_title": "About SCOCOEX",
        "contact_title": "Connect with SCOCOEX",
        "name": "Name",
        "company_org": "Company / Organization",
        "email": "Email",
        "interest": "Interest",
        "message": "Message",
        "send": "Send Inquiry",
        "received": "Thank you. Your inquiry has been received.",
        "gallery_title": "SCOCOEX Gallery",
        "gallery_text": "Event images can be added here later.",
        "no_docs": "No PDF or DOCX files were found inside docs/.",
        "docs_missing": "The docs folder does not exist. Create a folder named docs in the repository.",
        "no_companies": "No companies were loaded from companies.json.",
        "companies_error": "companies.json could not be loaded. The exact error is shown below.",
        "groq_missing": "Groq API is not configured yet. Add GROQ_API_KEY in Streamlit Secrets.",
        "no_results": "I could not find relevant information in the current Knowledge Base.",
        "no_text": "No readable text was found.",
        "download_pdf": "Download AI Answer as PDF",
        "source_file": "Source",
        "words": "Words",
        "numeric_values": "Numeric values",
        "extracted_sections": "Extracted sections",
    },
    "fa": {
        "nav": "منوی اصلی",
        "home": "خانه",
        "info": "مرکز اطلاعات",
        "assistant": "دستیار هوش مصنوعی",
        "companies": "شرکت‌های بین‌المللی",
        "intel": "داشبورد هوشمندی",
        "gallery": "گالری",
        "about": "درباره ما",
        "contact": "ارتباط با ما",
        "language": "زبان",
        "response_language": "زبان پاسخ هوش مصنوعی",
        "hero_label": "پلتفرم جهانی سرمایه‌گذاری و هوشمندی",
        "hero_title": "ارتباط.<br>کشف.<br>سرمایه‌گذاری.",
        "hero_text": "یک پلتفرم جهانی مبتنی بر هوش مصنوعی برای اتصال سازمان‌ها، سرمایه‌گذاران، شرکت‌ها و اطلاعات راهبردی.",
        "documents": "اسناد",
        "chunks": "قطعات دانش",
        "international_companies": "شرکت‌های بین‌المللی",
        "languages": "زبان‌ها",
        "knowledge": "دانش",
        "knowledge_text": "اسناد سازمانی، گزارش‌ها و اطلاعات راهبردی به یک پایگاه دانش قابل جستجو تبدیل می‌شوند.",
        "ai": "هوش مصنوعی",
        "ai_text": "هوش مصنوعی مبتنی بر Groq با استفاده از اطلاعات بازیابی‌شده از اسناد SCOCOEX پاسخ می‌دهد.",
        "network": "شبکه جهانی",
        "network_text": "شرکت‌ها، سرمایه‌گذاران و سازمان‌های راهبردی بین‌المللی را بررسی کنید.",
        "information_title": "مرکز اطلاعات",
        "information_text": "فایل‌های PDF و DOCX موجود در پوشه docs به‌صورت خودکار خوانده می‌شوند.",
        "search_documents": "جستجو در اسناد",
        "sections": "بخش استخراج‌شده",
        "summary": "ساخت خلاصه هوش مصنوعی",
        "summary_title": "خلاصه هوشمند سند",
        "ask_title": "دستیار هوشمند SCOCOEX",
        "ask_caption": "درباره اطلاعات موجود در پایگاه دانش SCOCOEX سؤال بپرسید.",
        "your_question": "سؤال شما",
        "ask_button": "پرسش از SCOCOEX AI",
        "searching": "در حال جستجو در پایگاه دانش SCOCOEX...",
        "generating": "در حال تولید تحلیل...",
        "sources": "منابع",
        "companies_title": "شرکت‌های بین‌المللی",
        "search_company": "جستجوی شرکت",
        "visit": "مشاهده وب‌سایت رسمی ←",
        "company_count": "سازمان",
        "intel_title": "داشبورد هوشمندی",
        "intel_caption": "لایه اولیه هوشمندی اسناد.",
        "document_size": "حجم اسناد",
        "numeric": "اطلاعات عددی",
        "about_title": "درباره SCOCOEX",
        "contact_title": "ارتباط با SCOCOEX",
        "name": "نام",
        "company_org": "شرکت / سازمان",
        "email": "ایمیل",
        "interest": "موضوع همکاری",
        "message": "پیام",
        "send": "ارسال درخواست",
        "received": "درخواست شما با موفقیت دریافت شد.",
        "gallery_title": "گالری SCOCOEX",
        "gallery_text": "تصاویر رویداد در این بخش قرار خواهند گرفت.",
        "no_docs": "هیچ فایل PDF یا DOCX در پوشه docs پیدا نشد.",
        "docs_missing": "پوشه docs وجود ندارد. یک پوشه با نام docs در Repository بسازید.",
        "no_companies": "هیچ شرکتی از companies.json بارگذاری نشد.",
        "companies_error": "فایل companies.json قابل بارگذاری نیست. خطای دقیق پایین نمایش داده شده است.",
        "groq_missing": "کلید Groq تنظیم نشده است. GROQ_API_KEY را در Streamlit Secrets قرار دهید.",
        "no_results": "اطلاعات مرتبطی در پایگاه دانش فعلی پیدا نشد.",
        "no_text": "متن قابل خواندن پیدا نشد.",
        "download_pdf": "دریافت پاسخ AI به صورت PDF",
        "source_file": "منبع",
        "words": "تعداد کلمات",
        "numeric_values": "مقادیر عددی",
        "extracted_sections": "بخش‌های استخراج‌شده",
    },
    "ar": {
        "nav": "التنقل",
        "home": "الرئيسية",
        "info": "مركز المعلومات",
        "assistant": "المساعد الذكي",
        "companies": "الشركات الدولية",
        "intel": "لوحة الذكاء",
        "gallery": "المعرض",
        "about": "من نحن",
        "contact": "تواصل معنا",
        "language": "اللغة",
        "response_language": "لغة رد الذكاء الاصطناعي",
        "hero_label": "منصة عالمية للاستثمار والذكاء",
        "hero_title": "تواصل.<br>اكتشف.<br>استثمر.",
        "hero_text": "منصة عالمية مدعومة بالذكاء الاصطناعي لربط المؤسسات والمستثمرين والشركات والمعلومات الاستراتيجية.",
        "documents": "الوثائق",
        "chunks": "مقاطع المعرفة",
        "international_companies": "الشركات الدولية",
        "languages": "اللغات",
        "knowledge": "المعرفة",
        "knowledge_text": "تحويل الوثائق والتقارير والمعلومات الاستراتيجية إلى قاعدة معرفة قابلة للبحث.",
        "ai": "الذكاء الاصطناعي",
        "ai_text": "يجيب الذكاء الاصطناعي المدعوم من Groq باستخدام المعلومات المسترجعة من وثائق SCOCOEX.",
        "network": "الشبكة العالمية",
        "network_text": "استكشف الشركات والمستثمرين والمنظمات الاستراتيجية الدولية.",
        "information_title": "مركز المعلومات",
        "information_text": "يتم فهرسة ملفات PDF وDOCX الموجودة في مجلد docs تلقائياً.",
        "search_documents": "البحث في الوثائق",
        "sections": "الأقسام المستخرجة",
        "summary": "إنشاء ملخص بالذكاء الاصطناعي",
        "summary_title": "ملخص ذكي للوثيقة",
        "ask_title": "مساعد SCOCOEX الذكي",
        "ask_caption": "اطرح أسئلة حول المعلومات الموجودة في قاعدة معرفة SCOCOEX.",
        "your_question": "سؤالك",
        "ask_button": "اسأل SCOCOEX AI",
        "searching": "جارٍ البحث في قاعدة المعرفة...",
        "generating": "جارٍ إنشاء التحليل...",
        "sources": "المصادر",
        "companies_title": "الشركات الدولية",
        "search_company": "البحث عن شركة",
        "visit": "زيارة الموقع الرسمي ←",
        "company_count": "منظمة",
        "intel_title": "لوحة الذكاء",
        "intel_caption": "طبقة أولية لتحليل الوثائق.",
        "document_size": "حجم الوثائق",
        "numeric": "المعلومات الرقمية",
        "about_title": "عن SCOCOEX",
        "contact_title": "تواصل مع SCOCOEX",
        "name": "الاسم",
        "company_org": "الشركة / المؤسسة",
        "email": "البريد الإلكتروني",
        "interest": "مجال الاهتمام",
        "message": "الرسالة",
        "send": "إرسال الطلب",
        "received": "تم استلام طلبك بنجاح.",
        "gallery_title": "معرض SCOCOEX",
        "gallery_text": "يمكن إضافة صور الفعالية هنا لاحقاً.",
        "no_docs": "لم يتم العثور على ملفات PDF أو DOCX داخل docs.",
        "docs_missing": "مجلد docs غير موجود. أنشئ مجلداً باسم docs في المستودع.",
        "no_companies": "لم يتم تحميل شركات من companies.json.",
        "companies_error": "تعذر تحميل companies.json. يظهر الخطأ الدقيق أدناه.",
        "groq_missing": "لم يتم إعداد Groq API بعد. أضف GROQ_API_KEY في Streamlit Secrets.",
        "no_results": "لم أجد معلومات ذات صلة في قاعدة المعرفة الحالية.",
        "no_text": "لم يتم العثور على نص قابل للقراءة.",
        "download_pdf": "تنزيل إجابة AI بصيغة PDF",
        "source_file": "المصدر",
        "words": "الكلمات",
        "numeric_values": "القيم الرقمية",
        "extracted_sections": "الأقسام المستخرجة",
    },
    "zh": {
        "nav": "导航",
        "home": "首页",
        "info": "信息中心",
        "assistant": "AI 助手",
        "companies": "国际企业",
        "intel": "智能分析",
        "gallery": "图库",
        "about": "关于我们",
        "contact": "联系我们",
        "language": "语言",
        "response_language": "AI 回复语言",
        "hero_label": "全球投资与智能平台",
        "hero_title": "连接。<br>探索。<br>投资。",
        "hero_text": "一个由人工智能驱动的全球知识与智能平台，用于连接组织、投资者、企业和战略信息。",
        "documents": "文件",
        "chunks": "知识片段",
        "international_companies": "国际企业",
        "languages": "语言",
        "knowledge": "知识",
        "knowledge_text": "将机构文件、报告和战略信息转化为可搜索的知识库。",
        "ai": "人工智能",
        "ai_text": "由 Groq 驱动的 AI 使用 SCOCOEX 文件中的检索信息回答问题。",
        "network": "全球网络",
        "network_text": "探索国际企业、投资者和战略组织。",
        "information_title": "信息中心",
        "information_text": "docs/ 文件夹中的 PDF 和 DOCX 会自动建立索引。",
        "search_documents": "搜索文件",
        "sections": "提取章节",
        "summary": "生成 AI 摘要",
        "summary_title": "AI 文件摘要",
        "ask_title": "SCOCOEX AI 助手",
        "ask_caption": "询问 SCOCOEX 知识库中的信息。",
        "your_question": "您的问题",
        "ask_button": "询问 SCOCOEX AI",
        "searching": "正在搜索知识库...",
        "generating": "正在生成分析...",
        "sources": "来源",
        "companies_title": "国际企业",
        "search_company": "搜索企业",
        "visit": "访问官方网站 →",
        "company_count": "家机构",
        "intel_title": "智能分析仪表板",
        "intel_caption": "基础文件智能分析层。",
        "document_size": "文件规模",
        "numeric": "数字信息",
        "about_title": "关于 SCOCOEX",
        "contact_title": "联系 SCOCOEX",
        "name": "姓名",
        "company_org": "公司 / 组织",
        "email": "邮箱",
        "interest": "合作方向",
        "message": "留言",
        "send": "发送请求",
        "received": "您的请求已收到。",
        "gallery_title": "SCOCOEX 图库",
        "gallery_text": "之后可以在这里加入活动图片。",
        "no_docs": "docs/ 中没有找到 PDF 或 DOCX 文件。",
        "docs_missing": "docs 文件夹不存在。请在 Repository 中创建 docs 文件夹。",
        "no_companies": "未从 companies.json 加载企业。",
        "companies_error": "无法加载 companies.json。下面显示具体错误。",
        "groq_missing": "尚未配置 Groq API。请在 Streamlit Secrets 中添加 GROQ_API_KEY。",
        "no_results": "当前知识库中没有找到相关信息。",
        "no_text": "没有找到可读取的文本。",
        "download_pdf": "下载 AI 回答 PDF",
        "source_file": "来源",
        "words": "词数",
        "numeric_values": "数字值",
        "extracted_sections": "提取章节",
    },
    "ru": {
        "nav": "НАВИГАЦИЯ",
        "home": "Главная",
        "info": "Информационный центр",
        "assistant": "AI-ассистент",
        "companies": "Международные компании",
        "intel": "Интеллект",
        "gallery": "Галерея",
        "about": "О нас",
        "contact": "Контакты",
        "language": "Язык",
        "response_language": "Язык ответа AI",
        "hero_label": "ГЛОБАЛЬНАЯ ПЛАТФОРМА ИНВЕСТИЦИЙ И АНАЛИТИКИ",
        "hero_title": "Связывай.<br>Исследуй.<br>Инвестируй.",
        "hero_text": "Глобальная интеллектуальная платформа на базе ИИ для объединения организаций, инвесторов, компаний и стратегической информации.",
        "documents": "Документы",
        "chunks": "Фрагменты знаний",
        "international_companies": "Международные компании",
        "languages": "Языки",
        "knowledge": "Знания",
        "knowledge_text": "Документы, отчёты и стратегическая информация превращаются в поисковую базу знаний.",
        "ai": "Искусственный интеллект",
        "ai_text": "AI на базе Groq отвечает на вопросы, используя информацию из документов SCOCOEX.",
        "network": "Глобальная сеть",
        "network_text": "Исследуйте международные компании, инвесторов и стратегические организации.",
        "information_title": "Информационный центр",
        "information_text": "PDF и DOCX в папке docs автоматически индексируются.",
        "search_documents": "Поиск документов",
        "sections": "Извлечённые разделы",
        "summary": "Создать AI-резюме",
        "summary_title": "AI-резюме документа",
        "ask_title": "AI-ассистент SCOCOEX",
        "ask_caption": "Задавайте вопросы по информации из базы знаний SCOCOEX.",
        "your_question": "Ваш вопрос",
        "ask_button": "Спросить SCOCOEX AI",
        "searching": "Поиск в базе знаний...",
        "generating": "Создание анализа...",
        "sources": "Источники",
        "companies_title": "Международные компании",
        "search_company": "Поиск компании",
        "visit": "Официальный сайт →",
        "company_count": "организаций",
        "intel_title": "Интеллектуальная панель",
        "intel_caption": "Базовый уровень анализа документов.",
        "document_size": "Размер документов",
        "numeric": "Числовая информация",
        "about_title": "О SCOCOEX",
        "contact_title": "Связаться с SCOCOEX",
        "name": "Имя",
        "company_org": "Компания / организация",
        "email": "Email",
        "interest": "Направление интереса",
        "message": "Сообщение",
        "send": "Отправить запрос",
        "received": "Ваш запрос получен.",
        "gallery_title": "Галерея SCOCOEX",
        "gallery_text": "Позже здесь можно добавить изображения мероприятия.",
        "no_docs": "В папке docs не найдено PDF или DOCX.",
        "docs_missing": "Папка docs не существует. Создайте папку docs в Repository.",
        "no_companies": "Компании из companies.json не загружены.",
        "companies_error": "Не удалось загрузить companies.json. Ниже показана точная ошибка.",
        "groq_missing": "Groq API ещё не настроен. Добавьте GROQ_API_KEY в Streamlit Secrets.",
        "no_results": "В текущей базе знаний не найдена релевантная информация.",
        "no_text": "Не найден читаемый текст.",
        "download_pdf": "Скачать ответ AI в PDF",
        "source_file": "Источник",
        "words": "Слова",
        "numeric_values": "Числовые значения",
        "extracted_sections": "Извлечённые разделы",
    },
}


# ============================================================
# UI CSS
# ============================================================

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: Inter, sans-serif;
}

[data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at 8% 0%, rgba(84, 76, 255, 0.16), transparent 28%),
        radial-gradient(circle at 95% 8%, rgba(0, 190, 255, 0.10), transparent 25%),
        #070A12;
}

[data-testid="stSidebar"] {
    background: rgba(7, 10, 18, 0.98);
    border-right: 1px solid rgba(255,255,255,0.07);
}

.block-container {
    max-width: 1450px;
    padding-top: 2.2rem;
}

.hero {
    padding: 58px 48px;
    border-radius: 32px;
    border: 1px solid rgba(255,255,255,0.10);
    background:
        radial-gradient(circle at 80% 20%, rgba(105, 91, 255, 0.18), transparent 32%),
        linear-gradient(135deg, rgba(27,31,57,0.97), rgba(9,12,23,0.96));
    box-shadow: 0 30px 90px rgba(0,0,0,0.28);
    margin-bottom: 30px;
}

.hero-label {
    color: #A6B0FF;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: 0.18em;
    text-transform: uppercase;
}

.hero h1 {
    font-size: clamp(46px, 7vw, 86px);
    line-height: 0.98;
    margin: 16px 0 22px 0;
    font-weight: 800;
}

.hero p {
    color: #B8C0D4;
    font-size: 17px;
    max-width: 850px;
    line-height: 1.8;
}

.metric-card, .feature-card, .company-card {
    padding: 24px;
    border-radius: 22px;
    background: rgba(255,255,255,0.045);
    border: 1px solid rgba(255,255,255,0.08);
    height: 100%;
    margin-bottom: 16px;
}

.metric-number {
    font-size: 30px;
    font-weight: 800;
    margin-top: 8px;
}

.muted {
    color: #8F99AE;
    font-size: 13px;
}

.answer-box {
    padding: 26px;
    border-radius: 22px;
    background: rgba(255,255,255,0.045);
    border: 1px solid rgba(255,255,255,0.08);
    line-height: 1.85;
}

.source-box {
    padding: 12px 16px;
    border-radius: 13px;
    background: rgba(105, 91, 255, 0.10);
    border: 1px solid rgba(105, 91, 255, 0.18);
    margin: 7px 0;
}

.company-logo {
    width: 52px;
    height: 52px;
    border-radius: 15px;
    object-fit: contain;
    background: rgba(255,255,255,0.92);
    padding: 6px;
}

.company-initial {
    width: 52px;
    height: 52px;
    border-radius: 15px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(255,255,255,0.08);
    font-size: 20px;
    font-weight: 800;
}

.section-title {
    font-size: 30px;
    font-weight: 800;
    margin-bottom: 8px;
}

.small-label {
    color: #929CAF;
    font-size: 12px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}

[dir="rtl"] .hero,
[dir="rtl"] .answer-box,
[dir="rtl"] .feature-card,
[dir="rtl"] .company-card,
[dir="rtl"] .metric-card {
    text-align: right;
}

[dir="rtl"] .hero-label {
    letter-spacing: normal;
}

@media (max-width: 700px) {
    .hero {
        padding: 38px 25px;
    }
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================

def t(key):
    return T[st.session_state.get("lang_code", "en")].get(key, key)


def escape(value):
    return html.escape(str(value or ""))


def get_direction():
    return "rtl" if st.session_state.get("lang_code") in {"fa", "ar"} else "ltr"


def apply_direction():
    direction = get_direction()
    st.markdown(
        f"""
<div dir="{direction}" style="display:none"></div>
<style>
.main .block-container {{
    direction: {direction};
}}
</style>
""",
        unsafe_allow_html=True,
    )


# ============================================================
# DOCUMENT EXTRACTION
# ============================================================

def extract_pdf(path):
    pages = []
    reader = PdfReader(str(path))

    for page_number, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""

        if text.strip():
            pages.append({"page": page_number, "text": text.strip()})

    return pages


def extract_docx(path):
    document = Document(str(path))
    sections = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            sections.append({"page": None, "text": text})

    for table_index, table in enumerate(document.tables, start=1):
        rows = []

        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows.append(" | ".join(cells))

        if rows:
            sections.append(
                {
                    "page": None,
                    "text": f"[TABLE {table_index}]\n" + "\n".join(rows),
                }
            )

    return sections


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
            sections = (
                extract_pdf(path)
                if path.suffix.lower() == ".pdf"
                else extract_docx(path)
            )

            documents.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
                    "sections": sections,
                    "error": None,
                }
            )
        except Exception as e:
            documents.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
                    "sections": [],
                    "error": str(e),
                }
            )

    return documents


def normalize_text(text):
    return re.sub(r"\s+", " ", text).strip()


@st.cache_data(show_spinner=False)
def create_chunks(documents, chunk_size=1400, overlap=200):
    chunks = []

    for document in documents:
        for section in document.get("sections", []):
            text = normalize_text(section.get("text", ""))

            if not text:
                continue

            start = 0
            step = max(1, chunk_size - overlap)

            while start < len(text):
                chunk_text = text[start:start + chunk_size]

                chunks.append(
                    {
                        "document": document["name"],
                        "page": section.get("page"),
                        "text": chunk_text,
                    }
                )

                start += step

    return chunks


# ============================================================
# LOCAL RETRIEVAL
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
        counts = Counter(chunk_tokens)

        score = sum(counts.get(token, 0) for token in question_tokens)

        # Small bonus for exact phrase matches
        if question.lower() in chunk["text"].lower():
            score += 8

        if score > 0:
            scored.append((score, chunk))

    scored.sort(key=lambda item: item[0], reverse=True)

    return [item[1] for item in scored[:top_k]]


# ============================================================
# GROQ
# ============================================================

def get_groq_client():
    api_key = None

    try:
        api_key = st.secrets.get("GROQ_API_KEY")
    except Exception:
        pass

    if not api_key:
        api_key = os.getenv("GROQ_API_KEY")

    if not api_key or Groq is None:
        return None

    return Groq(api_key=api_key)


def ask_groq(question, context, response_language):
    client = get_groq_client()

    if client is None:
        return None

    system_prompt = f"""
You are SCOCOEX NEXUS AI, an event intelligence and institutional
knowledge assistant.

Answer primarily from the supplied SCOCOEX document context.

Rules:
- Do not invent facts.
- If the supplied documents do not support an answer, say that clearly.
- Distinguish facts from reasonable interpretation.
- Preserve company names, project names and terminology.
- Answer in {response_language}.
- Be concise but useful for investors, delegates, companies and decision-makers.
- When relevant, identify the source document and page.
"""

    user_prompt = f"""
QUESTION:
{question}

DOCUMENT CONTEXT:
{context}

Return the best evidence-based answer.
"""

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.2,
            max_tokens=2200,
        )

        return response.choices[0].message.content

    except Exception as e:
        return f"Groq error: {e}"


def summarize_document(document):
    full_text = "\n".join(
        section["text"]
        for section in document.get("sections", [])
    )[:18000]

    if not full_text.strip():
        return t("no_text")

    client = get_groq_client()

    if client is None:
        return t("groq_missing")

    language_name = {
        "en": "English",
        "fa": "Persian",
        "ar": "Arabic",
        "zh": "Chinese",
        "ru": "Russian",
    }.get(st.session_state.get("lang_code", "en"), "English")

    prompt = f"""
Analyze this SCOCOEX document and produce a professional executive summary
in {language_name}.

Structure:
1. Executive Summary
2. Key Points
3. Organizations / Companies
4. Important Numbers
5. Opportunities
6. Strategic Takeaways

DOCUMENT:
{full_text}
"""

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "system",
                    "content": "You are a professional strategic intelligence analyst.",
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
            max_tokens=2400,
        )

        return response.choices[0].message.content

    except Exception as e:
        return f"Groq error: {e}"


# ============================================================
# PDF EXPORT
# ============================================================

def create_pdf(text, title="SCOCOEX NEXUS AI Report"):
    if SimpleDocTemplate is None:
        return None

    from io import BytesIO

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    story = [
        Paragraph(escape(title), styles["Title"]),
        Spacer(1, 16),
    ]

    for block in text.split("\n"):
        block = block.strip()
        if block:
            story.append(Paragraph(escape(block), styles["BodyText"]))
            story.append(Spacer(1, 7))

    doc.build(story)

    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# COMPANIES JSON
# ============================================================

@st.cache_data(show_spinner=False)
def load_companies():
    # First try the expected root location. If Streamlit/GitHub places the
    # file elsewhere, search the repository recursively.
    candidates = [COMPANIES_FILE]
    for candidate in ROOT.rglob("companies.json"):
        if candidate not in candidates:
            candidates.append(candidate)

    existing = next((p for p in candidates if p.is_file()), None)
    if existing is None:
        return [], "companies.json was not found in the deployed repository"

    try:
        raw = existing.read_text(encoding="utf-8-sig").strip()

        if not raw:
            return [], "File is empty"

        data = json.loads(raw)

        if isinstance(data, list):
            companies = data
        elif isinstance(data, dict):
            for key in ("companies", "data", "organizations", "items"):
                if isinstance(data.get(key), list):
                    companies = data[key]
                    break
            else:
                return [], "JSON object found, but no companies/data/organizations/items list exists"
        else:
            return [], "JSON root must be a list or an object containing a list"

        normalized = []

        for item in companies:
            if isinstance(item, str):
                normalized.append(
                    {
                        "name": item,
                        "website": "",
                        "sector": "",
                        "logo": "",
                    }
                )
                continue

            if not isinstance(item, dict):
                continue

            name = (
                item.get("name")
                or item.get("company")
                or item.get("title")
                or "International Organization"
            )

            website = (
                item.get("website")
                or item.get("url")
                or item.get("site")
                or ""
            )

            sector = (
                item.get("sector")
                or item.get("industry")
                or item.get("category")
                or "International Organization"
            )

            logo = item.get("logo") or item.get("logo_url") or ""

            normalized.append(
                {
                    "name": str(name),
                    "website": str(website),
                    "sector": str(sector),
                    "logo": str(logo),
                }
            )

        return normalized, None

    except Exception as e:
        return [], f"{type(e).__name__}: {e}"


# ============================================================
# SESSION STATE
# ============================================================

if "language_name" not in st.session_state:
    st.session_state.language_name = "English"

st.session_state.lang_code = LANGUAGES[st.session_state.language_name]
apply_direction()

documents = load_all_documents()
chunks = create_chunks(documents)
companies, companies_error = load_companies()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    """
<div style="padding:8px 0 24px 0;">
    <div style="font-size:27px;font-weight:800;">🌐 SCOCOEX</div>
    <div style="color:#8D98AF;font-size:13px;">NEXUS Intelligence Platform</div>
</div>
""",
    unsafe_allow_html=True,
)

st.sidebar.markdown(f"**{t('language')}**")

language_name = st.sidebar.selectbox(
    t("language"),
    list(LANGUAGES.keys()),
    index=list(LANGUAGES.keys()).index(st.session_state.language_name),
    label_visibility="collapsed",
)

if language_name != st.session_state.language_name:
    st.session_state.language_name = language_name
    st.rerun()

pages = [
    t("home"),
    t("info"),
    t("assistant"),
    t("companies"),
    t("intel"),
    t("gallery"),
    t("about"),
    t("contact"),
]

page = st.sidebar.radio(
    t("nav"),
    pages,
    label_visibility="collapsed",
)

st.sidebar.divider()

st.sidebar.caption(f"{t('documents')}: {len(documents)}")
st.sidebar.caption(f"{t('international_companies')}: {len(companies)}")


# ============================================================
# HOME
# ============================================================

if page == t("home"):

    # Official SCOCOEX logo
    if LOGO_PATH and LOGO_PATH.exists():
        logo_col = st.columns([1, 3, 1])[1]
        with logo_col:
            st.image(str(LOGO_PATH), use_container_width=True)

    st.markdown(
        f"""
<div class="hero" dir="{get_direction()}">
    <div class="hero-label">{t("hero_label")}</div>
    <h1>{t("hero_title")}</h1>
    <p>{t("hero_text")}</p>
</div>
""",
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)

    metrics = [
        (t("documents"), len(documents)),
        (t("chunks"), len(chunks)),
        (t("international_companies"), len(companies)),
        (t("languages"), "5"),
    ]

    for col, (label, value) in zip((c1, c2, c3, c4), metrics):
        with col:
            st.markdown(
                f"""
<div class="metric-card" dir="{get_direction()}">
    <div class="muted">{escape(label)}</div>
    <div class="metric-number">{escape(value)}</div>
</div>
""",
                unsafe_allow_html=True,
            )

    st.markdown(f"### {t('knowledge')}")

    col1, col2, col3 = st.columns(3)

    cards = [
        ("📚", t("knowledge"), t("knowledge_text")),
        ("🤖", t("ai"), t("ai_text")),
        ("🌍", t("network"), t("network_text")),
    ]

    for col, (icon, title, text) in zip((col1, col2, col3), cards):
        with col:
            st.markdown(
                f"""
<div class="feature-card" dir="{get_direction()}">
    <div style="font-size:28px;">{icon}</div>
    <h3>{escape(title)}</h3>
    <div class="muted" style="font-size:14px;line-height:1.7;">
        {escape(text)}
    </div>
</div>
""",
                unsafe_allow_html=True,
            )


# ============================================================
# INFORMATION CENTER
# ============================================================

elif page == t("info"):

    st.markdown(f'<div class="section-title">{t("information_title")}</div>', unsafe_allow_html=True)
    st.write(t("information_text"))

    if not DOCS_DIR.exists():
        st.error(t("docs_missing"))
    elif not documents:
        st.warning(t("no_docs"))
    else:
        st.success(f"{len(documents)} {t('documents').lower()}")

        search = st.text_input(t("search_documents"))

        filtered = [
            doc for doc in documents
            if not search or search.lower() in doc["name"].lower()
        ]

        for document in filtered:
            with st.expander(f"📄 {document['name']}"):

                col1, col2, col3 = st.columns(3)

                col1.metric(t("extracted_sections"), len(document["sections"]))

                total_words = sum(
                    len(tokenize(s["text"]))
                    for s in document["sections"]
                )

                col2.metric(t("words"), total_words)

                numeric_count = len(
                    re.findall(
                        r"\b\d+(?:[.,]\d+)?\b",
                        " ".join(s["text"] for s in document["sections"]),
                    )
                )

                col3.metric(t("numeric_values"), numeric_count)

                if document.get("error"):
                    st.error(document["error"])

                if st.button(
                    t("summary"),
                    key=f"summary_{document['name']}",
                    use_container_width=True,
                ):
                    with st.spinner(t("generating")):
                        summary = summarize_document(document)

                    st.markdown(
                        f'<div class="answer-box" dir="{get_direction()}">',
                        unsafe_allow_html=True,
                    )
                    st.markdown(summary)
                    st.markdown("</div>", unsafe_allow_html=True)

                    pdf_data = create_pdf(summary, f"SCOCOEX — {document['name']}")

                    if pdf_data:
                        st.download_button(
                            t("download_pdf"),
                            data=pdf_data,
                            file_name=f"SCOCOEX_{Path(document['name']).stem}_summary.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )


# ============================================================
# AI ASSISTANT
# ============================================================

elif page == t("assistant"):

    st.markdown(f'<div class="section-title">{t("ask_title")}</div>', unsafe_allow_html=True)
    st.write(t("ask_caption"))

    question = st.text_area(
        t("your_question"),
        height=140,
        placeholder="Ask about the SCOCOEX documents...",
    )

    if st.button(
        t("ask_button"),
        type="primary",
        use_container_width=True,
    ):

        if not question.strip():
            st.warning(t("your_question"))
        elif not chunks:
            st.warning(t("no_docs"))
        else:

            with st.spinner(t("searching")):
                results = retrieve_chunks(question, chunks, top_k=6)

            if not results:
                st.warning(t("no_results"))
            else:

                context_parts = []

                for result in results:
                    source = result["document"]
                    page_number = result.get("page")

                    source_info = (
                        f"{source} — Page {page_number}"
                        if page_number
                        else source
                    )

                    context_parts.append(
                        f"SOURCE: {source_info}\nCONTENT:\n{result['text']}"
                    )

                context = "\n\n".join(context_parts)

                language_name_map = {
                    "en": "English",
                    "fa": "Persian",
                    "ar": "Arabic",
                    "zh": "Chinese",
                    "ru": "Russian",
                }

                with st.spinner(t("generating")):
                    answer = ask_groq(
                        question,
                        context,
                        language_name_map[st.session_state.lang_code],
                    )

                if answer is None:
                    st.warning(t("groq_missing"))
                else:

                    st.markdown(
                        f'<div class="answer-box" dir="{get_direction()}">',
                        unsafe_allow_html=True,
                    )
                    st.markdown(answer)
                    st.markdown("</div>", unsafe_allow_html=True)

                    pdf_data = create_pdf(answer, "SCOCOEX NEXUS — AI Answer")

                    if pdf_data:
                        st.download_button(
                            t("download_pdf"),
                            data=pdf_data,
                            file_name="SCOCOEX_NEXUS_AI_Answer.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )

                    st.markdown(f"### {t('sources')}")

                    seen = set()

                    for result in results:
                        key = (result["document"], result.get("page"))

                        if key in seen:
                            continue

                        seen.add(key)

                        source = result["document"]
                        page_number = result.get("page")

                        source_text = (
                            f"📄 {source} — Page {page_number}"
                            if page_number
                            else f"📄 {source}"
                        )

                        st.markdown(
                            f'<div class="source-box" dir="{get_direction()}">{escape(source_text)}</div>',
                            unsafe_allow_html=True,
                        )


# ============================================================
# COMPANIES
# ============================================================

elif page == t("companies"):

    st.markdown(f'<div class="section-title">{t("companies_title")}</div>', unsafe_allow_html=True)

    if companies_error:
        st.error(t("companies_error"))
        st.code(companies_error)
        st.info(f"Expected file path: {COMPANIES_FILE}")

    elif not companies:
        st.warning(t("no_companies"))
        st.info(f"Expected file path: {COMPANIES_FILE}")

    else:

        search = st.text_input(t("search_company"))

        filtered = [
            company for company in companies
            if not search
            or search.lower() in json.dumps(
                company,
                ensure_ascii=False
            ).lower()
        ]

        st.caption(f"{len(filtered)} {t('company_count')}")

        columns = st.columns(3)

        for index, company in enumerate(filtered):

            with columns[index % 3]:

                name = company.get("name", "International Organization")
                website = company.get("website", "")
                sector = company.get("sector", "")
                logo = company.get("logo", "")

                if not logo and website:
                    domain = re.sub(r"^https?://", "", website).split("/")[0]
                    logo = f"https://www.google.com/s2/favicons?domain={domain}&sz=128"

                logo_html = ""

                if logo:
                    logo_html = (
                        f'<img class="company-logo" src="{escape(logo)}" '
                        f'alt="{escape(name)}">'
                    )
                else:
                    initial = escape(name[:1].upper())
                    logo_html = f'<div class="company-initial">{initial}</div>'

                link_html = ""

                if website:
                    safe_website = escape(website)
                    link_html = (
                        f'<a href="{safe_website}" target="_blank" '
                        f'rel="noopener noreferrer">{escape(t("visit"))}</a>'
                    )

                st.markdown(
                    f"""
<div class="company-card" dir="{get_direction()}">
    <div style="display:flex;gap:14px;align-items:center;">
        {logo_html}
        <div>
            <div style="font-size:18px;font-weight:700;">
                {escape(name)}
            </div>
            <div class="muted">
                {escape(sector)}
            </div>
        </div>
    </div>
    <div style="margin-top:20px;">
        {link_html}
    </div>
</div>
""",
                    unsafe_allow_html=True,
                )


# ============================================================
# INTELLIGENCE
# ============================================================

elif page == t("intel"):

    st.markdown(f'<div class="section-title">{t("intel_title")}</div>', unsafe_allow_html=True)
    st.write(t("intel_caption"))

    if not documents:
        st.warning(t("no_docs"))
    else:

        rows = []

        for document in documents:
            full_text = " ".join(
                section["text"]
                for section in document["sections"]
            )

            rows.append(
                {
                    t("source_file"): document["name"],
                    t("words"): len(tokenize(full_text)),
                    t("numeric_values"): len(
                        re.findall(r"\b\d+(?:[.,]\d+)?\b", full_text)
                    ),
                    t("extracted_sections"): len(document["sections"]),
                }
            )

        df = pd.DataFrame(rows)

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )

        st.markdown(f"### {t('document_size')}")

        st.bar_chart(
            df.set_index(t("source_file"))[t("words")]
        )

        st.markdown(f"### {t('numeric')}")

        st.bar_chart(
            df.set_index(t("source_file"))[t("numeric_values")]
        )


# ============================================================
# GALLERY
# ============================================================

elif page == t("gallery"):

    st.markdown(f'<div class="section-title">{t("gallery_title")}</div>', unsafe_allow_html=True)
    st.write(t("gallery_text"))


# ============================================================
# ABOUT
# ============================================================

elif page == t("about"):

    st.markdown(f'<div class="section-title">{t("about_title")}</div>', unsafe_allow_html=True)

    st.markdown(
        f"""
<div class="feature-card" dir="{get_direction()}">
    <h3>SCOCOEX NEXUS</h3>
    <p>
    {escape(t("hero_text"))}
    </p>
</div>
""",
        unsafe_allow_html=True,
    )


# ============================================================
# CONTACT
# ============================================================

elif page == t("contact"):

    st.markdown(f'<div class="section-title">{t("contact_title")}</div>', unsafe_allow_html=True)

    with st.form("contact_form"):

        name = st.text_input(t("name"))
        company = st.text_input(t("company_org"))
        email = st.text_input(t("email"))

        interest = st.selectbox(
            t("interest"),
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
            t("message"),
            height=160,
        )

        submitted = st.form_submit_button(
            t("send"),
            type="primary",
            use_container_width=True,
        )

        if submitted:
            st.success(t("received"))
