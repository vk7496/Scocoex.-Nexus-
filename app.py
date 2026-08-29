import os
import re
import json
import html
import textwrap
from pathlib import Path
from collections import Counter
from io import BytesIO
from urllib.request import Request, urlopen

import streamlit as st
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
    Paragraph = None
    Spacer = None
    getSampleStyleSheet = None


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="SCOCOEX NEXUS",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent

DOCS_DIR = ROOT / "docs"
ASSETS_DIR = ROOT / "assets"

COMPANIES_FILE = ROOT / "companies.json"

LOGO_FILE = ASSETS_DIR / "scocoex_logo-1.png"

SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".docx",
}

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

        "hero_label":
            "GLOBAL INVESTMENT & INTELLIGENCE PLATFORM",

        "hero_title":
            "Connect.<br>Discover.<br>Invest.",

        "hero_text":
            "An AI-powered global intelligence and knowledge platform designed to connect organizations, investors, companies and strategic information.",

        "documents": "Documents",
        "chunks": "Knowledge Chunks",
        "international_companies": "International Companies",
        "languages": "Languages",

        "knowledge": "Knowledge",
        "knowledge_text":
            "Institutional documents, reports and strategic information become a searchable knowledge base.",

        "ai": "Artificial Intelligence",
        "ai_text":
            "Groq-powered AI answers questions using information retrieved from SCOCOEX documents.",

        "network": "Global Network",
        "network_text":
            "Explore international companies, investors and strategic organizations.",

        "information_title": "Information Center",
        "information_text":
            "PDF and DOCX files inside the docs/ folder are automatically indexed.",

        "search_documents": "Search documents",
        "summary": "Generate AI Summary",
        "summary_title": "AI Document Summary",

        "ask_title": "SCOCOEX AI Assistant",
        "ask_caption":
            "Ask questions about the information contained in the SCOCOEX Knowledge Base.",

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
        "intel_caption":
            "Document intelligence and numerical overview.",

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
        "gallery_text":
            "Event images and global activity can be presented here.",

        "no_docs":
            "No PDF or DOCX files were found inside docs/.",

        "docs_missing":
            "The docs folder does not exist.",

        "no_companies":
            "No companies were loaded.",

        "companies_error":
            "companies.json could not be loaded.",

        "groq_missing":
            "Groq API is not configured. Add GROQ_API_KEY in Streamlit Secrets.",

        "no_results":
            "I could not find relevant information in the current Knowledge Base.",

        "no_text":
            "No readable text was found.",

        "download_pdf":
            "Download AI Answer as PDF",

        "source_file": "Source",
        "words": "Words",
        "numeric_values": "Numeric values",
        "extracted_sections": "Extracted sections",
        "sector": "Sector",
        "company_description": "Description",
        "country": "Country",
        "all": "All",
        "all_companies": "All organizations",
        "strategic": "Strategic Organizations",
        "international_network": "International Company Network",
        "search_network":
            "Search organizations, sectors or countries...",
        "organizations_in_network":
            "organizations in network",
        "official": "Official Website",
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

        "hero_label":
            "پلتفرم جهانی سرمایه‌گذاری و هوشمندی",

        "hero_title":
            "ارتباط.<br>کشف.<br>سرمایه‌گذاری.",

        "hero_text":
            "یک پلتفرم جهانی مبتنی بر هوش مصنوعی برای اتصال سازمان‌ها، سرمایه‌گذاران، شرکت‌ها و اطلاعات راهبردی.",

        "documents": "اسناد",
        "chunks": "قطعات دانش",
        "international_companies": "شرکت‌های بین‌المللی",
        "languages": "زبان‌ها",

        "knowledge": "دانش",
        "knowledge_text":
            "اسناد سازمانی، گزارش‌ها و اطلاعات راهبردی به یک پایگاه دانش قابل جستجو تبدیل می‌شوند.",

        "ai": "هوش مصنوعی",
        "ai_text":
            "هوش مصنوعی مبتنی بر Groq با استفاده از اطلاعات بازیابی‌شده از اسناد SCOCOEX پاسخ می‌دهد.",

        "network": "شبکه جهانی",
        "network_text":
            "شرکت‌ها، سرمایه‌گذاران و سازمان‌های راهبردی بین‌المللی را بررسی کنید.",

        "information_title": "مرکز اطلاعات",
        "information_text":
            "فایل‌های PDF و DOCX موجود در پوشه docs به‌صورت خودکار خوانده و ایندکس می‌شوند.",

        "search_documents": "جستجو در اسناد",
        "summary": "ساخت خلاصه هوش مصنوعی",
        "summary_title": "خلاصه هوشمند سند",

        "ask_title": "دستیار هوشمند SCOCOEX",
        "ask_caption":
            "درباره اطلاعات موجود در پایگاه دانش SCOCOEX سؤال بپرسید.",

        "your_question": "سؤال شما",
        "ask_button": "پرسش از SCOCOEX AI",

        "searching":
            "در حال جستجو در پایگاه دانش SCOCOEX...",

        "generating":
            "در حال تولید تحلیل...",

        "sources": "منابع",

        "companies_title": "شرکت‌های بین‌المللی",
        "search_company": "جستجوی شرکت",
        "visit": "مشاهده وب‌سایت رسمی ←",

        "company_count": "سازمان",

        "intel_title": "داشبورد هوشمندی",
        "intel_caption":
            "تحلیل اسناد و مرور اطلاعات عددی.",

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

        "received":
            "درخواست شما با موفقیت دریافت شد.",

        "gallery_title": "گالری SCOCOEX",
        "gallery_text":
            "تصاویر رویدادها و فعالیت‌های جهانی را می‌توان در این بخش نمایش داد.",

        "no_docs":
            "هیچ فایل PDF یا DOCX در پوشه docs پیدا نشد.",

        "docs_missing":
            "پوشه docs وجود ندارد.",

        "no_companies":
            "هیچ شرکتی بارگذاری نشد.",

        "companies_error":
            "فایل companies.json قابل بارگذاری نیست.",

        "groq_missing":
            "کلید Groq تنظیم نشده است. GROQ_API_KEY را در Streamlit Secrets قرار دهید.",

        "no_results":
            "اطلاعات مرتبطی در پایگاه دانش فعلی پیدا نشد.",

        "no_text":
            "متن قابل خواندن پیدا نشد.",

        "download_pdf":
            "دریافت پاسخ AI به صورت PDF",

        "source_file": "منبع",
        "words": "کلمات",
        "numeric_values": "مقادیر عددی",
        "extracted_sections": "بخش‌های استخراج‌شده",
        "sector": "حوزه",
        "company_description": "توضیحات",
        "country": "کشور",
        "all": "همه",
        "all_companies": "همه سازمان‌ها",
        "strategic": "سازمان‌های راهبردی",
        "international_network":
            "شبکه شرکت‌های بین‌المللی",
        "search_network":
            "جستجوی سازمان، حوزه یا کشور...",
        "organizations_in_network":
            "سازمان در شبکه",
        "official":
            "وب‌سایت رسمی",
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

        "hero_label":
            "منصة عالمية للاستثمار والذكاء",

        "hero_title":
            "تواصل.<br>اكتشف.<br>استثمر.",

        "hero_text":
            "منصة عالمية مدعومة بالذكاء الاصطناعي لربط المؤسسات والمستثمرين والشركات والمعلومات الاستراتيجية.",

        "documents": "الوثائق",
        "chunks": "مقاطع المعرفة",
        "international_companies": "الشركات الدولية",
        "languages": "اللغات",

        "knowledge": "المعرفة",
        "knowledge_text":
            "تحويل الوثائق والتقارير والمعلومات الاستراتيجية إلى قاعدة معرفة قابلة للبحث.",

        "ai": "الذكاء الاصطناعي",
        "ai_text":
            "يجيب الذكاء الاصطناعي المدعوم من Groq باستخدام المعلومات المسترجعة من وثائق SCOCOEX.",

        "network": "الشبكة العالمية",
        "network_text":
            "استكشف الشركات والمستثمرين والمنظمات الاستراتيجية الدولية.",

        "information_title": "مركز المعلومات",
        "information_text":
            "يتم فهرسة ملفات PDF وDOCX الموجودة في مجلد docs تلقائياً.",

        "search_documents": "البحث في الوثائق",
        "summary": "إنشاء ملخص بالذكاء الاصطناعي",
        "summary_title": "ملخص ذكي للوثيقة",

        "ask_title": "مساعد SCOCOEX الذكي",
        "ask_caption":
            "اطرح أسئلة حول المعلومات الموجودة في قاعدة معرفة SCOCOEX.",

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
        "intel_caption":
            "تحليل الوثائق ومراجعة المعلومات الرقمية.",

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
        "gallery_text":
            "يمكن عرض صور الفعاليات والأنشطة العالمية هنا.",

        "no_docs":
            "لم يتم العثور على ملفات PDF أو DOCX داخل docs.",

        "docs_missing":
            "مجلد docs غير موجود.",

        "no_companies":
            "لم يتم تحميل شركات.",

        "companies_error":
            "تعذر تحميل companies.json.",

        "groq_missing":
            "لم يتم إعداد Groq API. أضف GROQ_API_KEY في Streamlit Secrets.",

        "no_results":
            "لم أجد معلومات ذات صلة في قاعدة المعرفة الحالية.",

        "no_text":
            "لم يتم العثور على نص قابل للقراءة.",

        "download_pdf":
            "تنزيل إجابة AI بصيغة PDF",

        "source_file": "المصدر",
        "words": "الكلمات",
        "numeric_values": "القيم الرقمية",
        "extracted_sections": "الأقسام المستخرجة",
        "sector": "القطاع",
        "company_description": "الوصف",
        "country": "الدولة",
        "all": "الكل",
        "all_companies": "جميع المؤسسات",
        "strategic": "المنظمات الاستراتيجية",
        "international_network":
            "شبكة الشركات الدولية",
        "search_network":
            "البحث عن المؤسسات أو القطاعات أو الدول...",
        "organizations_in_network":
            "منظمة في الشبكة",
        "official":
            "الموقع الرسمي",
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
        "hero_text":
            "一个由人工智能驱动的全球知识与智能平台，用于连接组织、投资者、企业和战略信息。",

        "documents": "文件",
        "chunks": "知识片段",
        "international_companies": "国际企业",
        "languages": "语言",

        "knowledge": "知识",
        "knowledge_text":
            "将机构文件、报告和战略信息转化为可搜索的知识库。",

        "ai": "人工智能",
        "ai_text":
            "由 Groq 驱动的 AI 使用 SCOCOEX 文件中的检索信息回答问题。",

        "network": "全球网络",
        "network_text":
            "探索国际企业、投资者和战略组织。",

        "information_title": "信息中心",
        "information_text":
            "docs/ 文件夹中的 PDF 和 DOCX 会自动建立索引。",

        "search_documents": "搜索文件",
        "summary": "生成 AI 摘要",
        "summary_title": "AI 文件摘要",

        "ask_title": "SCOCOEX AI 助手",
        "ask_caption":
            "询问 SCOCOEX 知识库中的信息。",
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
        "intel_caption":
            "文件智能分析和数字信息概览。",

        "document_size": "文件规模",
        "numeric": "数字信息",

        "about_title": "关于 SCOCOEX",
        "contact_title": "联系我们",

        "name": "姓名",
        "company_org": "公司 / 组织",
        "email": "邮箱",
        "interest": "合作方向",
        "message": "留言",
        "send": "发送请求",

        "received": "您的请求已收到。",

        "gallery_title": "SCOCOEX 图库",
        "gallery_text":
            "之后可以在这里展示活动和全球业务图片。",

        "no_docs": "docs/ 中没有找到 PDF 或 DOCX 文件。",
        "docs_missing": "docs 文件夹不存在。",
        "no_companies": "未加载企业。",
        "companies_error": "无法加载 companies.json。",
        "groq_missing":
            "尚未配置 Groq API。请在 Streamlit Secrets 中添加 GROQ_API_KEY。",
        "no_results": "当前知识库中没有找到相关信息。",
        "no_text": "没有找到可读取的文本。",
        "download_pdf": "下载 AI 回答 PDF",

        "source_file": "来源",
        "words": "词数",
        "numeric_values": "数字值",
        "extracted_sections": "提取章节",
        "sector": "领域",
        "company_description": "描述",
        "country": "国家",
        "all": "全部",
        "all_companies": "所有机构",
        "strategic": "战略组织",
        "international_network": "国际企业网络",
        "search_network":
            "搜索组织、领域或国家...",
        "organizations_in_network":
            "网络中的组织",
        "official": "官方网站",
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

        "hero_label":
            "ГЛОБАЛЬНАЯ ПЛАТФОРМА ИНВЕСТИЦИЙ И АНАЛИТИКИ",

        "hero_title":
            "Связывай.<br>Исследуй.<br>Инвестируй.",

        "hero_text":
            "Глобальная интеллектуальная платформа на базе ИИ для объединения организаций, инвесторов, компаний и стратегической информации.",

        "documents": "Документы",
        "chunks": "Фрагменты знаний",
        "international_companies": "Международные компании",
        "languages": "Языки",

        "knowledge": "Знания",
        "knowledge_text":
            "Документы, отчёты и стратегическая информация превращаются в поисковую базу знаний.",

        "ai": "Искусственный интеллект",
        "ai_text":
            "AI на базе Groq отвечает на вопросы, используя информацию из документов SCOCOEX.",

        "network": "Глобальная сеть",
        "network_text":
            "Исследуйте международные компании, инвесторов и стратегические организации.",

        "information_title": "Информационный центр",
        "information_text":
            "PDF и DOCX в папке docs автоматически индексируются.",

        "search_documents": "Поиск документов",
        "summary": "Создать AI-резюме",
        "summary_title": "AI-резюме документа",

        "ask_title": "AI-ассистент SCOCOEX",
        "ask_caption":
            "Задавайте вопросы по информации из базы знаний SCOCOEX.",
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
        "intel_caption":
            "Анализ документов и обзор числовой информации.",

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
        "gallery_text":
            "Здесь можно показывать изображения мероприятий и глобальной деятельности.",

        "no_docs":
            "В папке docs не найдено PDF или DOCX.",

        "docs_missing":
            "Папка docs не существует.",

        "no_companies":
            "Компании не загружены.",

        "companies_error":
            "Не удалось загрузить companies.json.",

        "groq_missing":
            "Groq API ещё не настроен. Добавьте GROQ_API_KEY в Streamlit Secrets.",

        "no_results":
            "В текущей базе знаний не найдена релевантная информация.",

        "no_text":
            "Не найден читаемый текст.",

        "download_pdf":
            "Скачать ответ AI в PDF",

        "source_file": "Источник",
        "words": "Слова",
        "numeric_values": "Числовые значения",
        "extracted_sections": "Извлечённые разделы",
        "sector": "Сектор",
        "company_description": "Описание",
        "country": "Страна",
        "all": "Все",
        "all_companies": "Все организации",
        "strategic": "Стратегические организации",
        "international_network":
            "Международная сеть компаний",
        "search_network":
            "Поиск организаций, отраслей или стран...",
        "organizations_in_network":
            "организаций в сети",
        "official":
            "Официальный сайт",
    },
}


# ============================================================
# SESSION
# ============================================================

if "language_name" not in st.session_state:
    st.session_state.language_name = "English"

if "lang_code" not in st.session_state:
    st.session_state.lang_code = "en"


def t(key):
    lang = st.session_state.get("lang_code", "en")
    return T.get(lang, T["en"]).get(key, key)


def esc(value):
    return html.escape(str(value or ""))


def direction():
    if st.session_state.get("lang_code") in {"fa", "ar"}:
        return "rtl"
    return "ltr"


# ============================================================
# IMPORTANT HTML RENDERER
# ============================================================
#
# This solves the previous problem where HTML appeared literally
# on screen as:
# <div>
# <h1>
# </div>
#
# because indentation caused Markdown to interpret it as code.
# ============================================================

def render_html(content):
    """
    Render raw HTML directly instead of passing it through
    the Markdown parser.

    This prevents Streamlit from displaying HTML tags such as
    <div>, <h1>, <p>, etc. as literal code/text.
    """
    cleaned = textwrap.dedent(str(content)).strip()

    if not cleaned:
        return

    try:
        st.html(cleaned)
    except AttributeError:
        # Compatibility fallback for older Streamlit versions.
        st.markdown(
            cleaned,
            unsafe_allow_html=True,
        )


# ============================================================
# CSS
# ============================================================

render_html(
    """
    <style>

    @import url(
        'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap'
    );

    html, body, [class*="css"] {
        font-family: Inter, sans-serif;
    }

    [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(
                circle at 8% 0%,
                rgba(84,76,255,.18),
                transparent 28%
            ),
            radial-gradient(
                circle at 95% 8%,
                rgba(0,190,255,.10),
                transparent 25%
            ),
            #070A12;
    }

    [data-testid="stHeader"] {
        background: transparent !important;
    }

    .block-container {
        max-width: 1450px;
        padding-top: 2rem;
        padding-bottom: 5rem;
    }


    /* ================= SIDEBAR ================= */

    [data-testid="stSidebar"] {
        background:
            radial-gradient(
                circle at 20% 0%,
                rgba(102,92,255,.20),
                transparent 32%
            ),
            linear-gradient(
                180deg,
                #080B16 0%,
                #0D1222 55%,
                #080B14 100%
            ) !important;

        border-right:
            1px solid rgba(255,255,255,.09) !important;
    }

    [data-testid="stSidebar"] * {
        color: #F4F6FC !important;
    }

    [data-testid="stSidebar"] label {
        color: #D5DBEA !important;
        font-weight: 600 !important;
    }

    [data-testid="stSidebar"] [data-baseweb="select"] {
        background:
            rgba(255,255,255,.07) !important;

        border:
            1px solid rgba(255,255,255,.15) !important;

        border-radius:
            12px !important;
    }

    [data-testid="stSidebar"]
    [data-baseweb="select"] * {
        color: #FFFFFF !important;
    }

    [data-testid="stSidebar"]
    [data-testid="stRadio"] label {
        color: #E9ECF7 !important;
        padding: 9px 11px !important;
        border-radius: 12px !important;
        margin-bottom: 3px !important;
        transition: .18s ease;
    }

    [data-testid="stSidebar"]
    [data-testid="stRadio"] label:hover {
        background:
            rgba(145,132,255,.15) !important;
    }


    /* ================= HERO ================= */

    .hero {
        padding: 58px 48px;
        border-radius: 32px;

        border:
            1px solid rgba(255,255,255,.10);

        background:
            radial-gradient(
                circle at 82% 20%,
                rgba(105,91,255,.24),
                transparent 34%
            ),
            radial-gradient(
                circle at 20% 90%,
                rgba(0,190,255,.08),
                transparent 30%
            ),
            linear-gradient(
                135deg,
                rgba(27,31,57,.98),
                rgba(9,12,23,.97)
            );

        box-shadow:
            0 30px 90px rgba(0,0,0,.32);

        margin-bottom: 30px;
    }

    .hero-label {
        color: #AEB7FF;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: .18em;
        text-transform: uppercase;
    }

    .hero h1 {
        font-size: clamp(46px, 7vw, 86px);
        line-height: .98;
        margin: 18px 0 22px;
        font-weight: 800;
        color: #F7F8FC;
    }

    .hero p {
        color: #B8C0D4;
        font-size: 17px;
        max-width: 900px;
        line-height: 1.8;
    }


    /* ================= LOGO ================= */

    .logo-wrap {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 5px auto 30px auto;
        padding: 8px;
    }


    /* ================= METRICS ================= */

    .metric-card {
        padding: 24px;
        border-radius: 22px;

        background:
            linear-gradient(
                145deg,
                rgba(255,255,255,.075),
                rgba(255,255,255,.025)
            );

        border:
            1px solid rgba(255,255,255,.09);

        min-height: 120px;

        box-shadow:
            0 16px 40px rgba(0,0,0,.16);
    }

    .metric-label {
        color: #9EA8BD;
        font-size: 13px;
        font-weight: 600;
    }

    .metric-number {
        font-size: 32px;
        font-weight: 800;
        color: #F6F7FB;
        margin-top: 9px;
    }


    /* ================= FEATURE CARDS ================= */

    .feature-card {
        padding: 27px;
        border-radius: 24px;

        background:
            linear-gradient(
                145deg,
                rgba(255,255,255,.065),
                rgba(255,255,255,.025)
            );

        border:
            1px solid rgba(255,255,255,.08);

        height: 100%;
        min-height: 220px;

        box-shadow:
            0 18px 45px rgba(0,0,0,.16);
    }

    .feature-icon {
        font-size: 31px;
        margin-bottom: 14px;
    }

    .feature-title {
        font-size: 21px;
        font-weight: 800;
        color: #F2F4FA;
        margin-bottom: 10px;
    }

    .feature-text {
        color: #9EA8BD;
        line-height: 1.75;
        font-size: 14px;
    }


    /* ================= SECTION ================= */

    .section-title {
        font-size: 32px;
        font-weight: 800;
        color: #F4F6FB;
        margin-bottom: 7px;
    }

    .section-subtitle {
        color: #929CB2;
        font-size: 14px;
        margin-bottom: 25px;
    }


    /* ================= ANSWER ================= */

    .answer-box {
        padding: 28px;
        border-radius: 23px;

        background:
            linear-gradient(
                145deg,
                rgba(255,255,255,.065),
                rgba(255,255,255,.025)
            );

        border:
            1px solid rgba(255,255,255,.09);

        color: #E9ECF5;
        line-height: 1.85;

        box-shadow:
            0 18px 50px rgba(0,0,0,.16);
    }


    /* ================= COMPANY DIRECTORY ================= */

    .directory-header {
        font-size: 12px;
        font-weight: 800;
        letter-spacing: .14em;
        color: #AEB7FF;
        margin: 35px 0 17px;
    }

    .directory-header span {
        color: #C9A86A;
        margin-right: 7px;
    }

    .company-card {
        padding: 23px;
        border-radius: 22px;

        background:
            linear-gradient(
                145deg,
                rgba(255,255,255,.065),
                rgba(255,255,255,.025)
            );

        border:
            1px solid rgba(255,255,255,.08);

        min-height: 220px;

        box-shadow:
            0 15px 45px rgba(0,0,0,.16);

        margin-bottom: 18px;

        transition:
            transform .18s ease,
            border-color .18s ease;
    }

    .company-card:hover {
        transform: translateY(-4px);
        border-color:
            rgba(166,176,255,.30);
    }

    .company-name {
        color: #F4F6FB;
        font-size: 20px;
        font-weight: 800;
        margin-top: 12px;
    }

    .company-sector {
        color: #A9B2C8;
        font-size: 13px;
        margin-top: 7px;
    }

    .company-country {
        color: #8792AA;
        font-size: 12px;
        margin-top: 5px;
    }

    .company-description {
        color: #9EA8BD;
        font-size: 13px;
        line-height: 1.65;
        margin-top: 14px;
        min-height: 43px;
    }

    .company-initial {
        width: 62px;
        height: 62px;
        border-radius: 17px;

        display: flex;
        align-items: center;
        justify-content: center;

        background:
            linear-gradient(
                135deg,
                #202743,
                #111626
            );

        border:
            1px solid rgba(255,255,255,.10);

        color: #DDE1FF;
        font-size: 23px;
        font-weight: 800;
    }

    .company-link {
        display: inline-block;
        margin-top: 17px;

        color: #B9C0FF !important;
        text-decoration: none !important;
        font-size: 13px;
        font-weight: 700;
    }


    /* ================= STRATEGIC ================= */

    .strategic-card {
        padding: 28px;
        border-radius: 26px;
        min-height: 285px;

        background:
            linear-gradient(
                145deg,
                rgba(255,255,255,.085),
                rgba(255,255,255,.025)
            );

        border:
            1px solid rgba(255,255,255,.11);

        box-shadow:
            0 24px 60px rgba(0,0,0,.22);
    }

    .strategic-badge {
        color: #C9A86A;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: .14em;
        text-transform: uppercase;
        margin-bottom: 13px;
    }

    .strategic-name {
        color: #FFFFFF;
        font-size: 24px;
        font-weight: 800;
        margin-bottom: 10px;
    }

    .strategic-sector {
        color: #AEB7FF;
        font-size: 13px;
        margin-bottom: 18px;
    }

    .strategic-copy {
        color: #A2ACC0;
        font-size: 13px;
        line-height: 1.7;
        min-height: 70px;
    }

    .strategic-link {
        display: inline-block;
        margin-top: 18px;
        color: #D5C18A !important;
        font-weight: 700;
        text-decoration: none !important;
    }


    /* ================= RTL ================= */

    [dir="rtl"] {
        text-align: right;
    }


    /* ================= MOBILE ================= */

    @media (max-width: 700px) {

        .block-container {
            padding-top: 1rem;
        }

        .hero {
            padding: 38px 25px;
        }

        .hero h1 {
            font-size: 45px;
        }

        .hero p {
            font-size: 15px;
        }

        .section-title {
            font-size: 27px;
        }

    }

    </style>
    """
)


# ============================================================
# LANGUAGE SELECTOR
# ============================================================
#
# EXACTLY ONE language selectbox.
# This prevents StreamlitDuplicateElementId.
# ============================================================

selected_language = st.sidebar.selectbox(
    t("language"),
    list(LANGUAGES.keys()),
    index=list(LANGUAGES.keys()).index(
        st.session_state.language_name
    ),
    key="global_language_selector_v5",
)

if selected_language != st.session_state.language_name:

    st.session_state.language_name = selected_language
    st.session_state.lang_code = LANGUAGES[selected_language]

    st.rerun()


# ============================================================
# DOCUMENT EXTRACTION
# ============================================================

def extract_pdf(path):

    sections = []

    try:
        reader = PdfReader(str(path))

        for page_no, page in enumerate(
            reader.pages,
            start=1,
        ):

            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""

            if text.strip():

                sections.append(
                    {
                        "page": page_no,
                        "text": text.strip(),
                    }
                )

    except Exception:
        return []

    return sections


def extract_docx(path):

    sections = []

    try:
        document = Document(str(path))

        for paragraph in document.paragraphs:

            text = paragraph.text.strip()

            if text:
                sections.append(
                    {
                        "page": None,
                        "text": text,
                    }
                )

        for table_no, table in enumerate(
            document.tables,
            start=1,
        ):

            rows = []

            for row in table.rows:

                cells = [
                    cell.text.strip()
                    for cell in row.cells
                ]

                rows.append(
                    " | ".join(cells)
                )

            if rows:

                sections.append(
                    {
                        "page": None,
                        "text":
                            f"[TABLE {table_no}]\n"
                            + "\n".join(rows),
                    }
                )

    except Exception:
        return []

    return sections


@st.cache_data(show_spinner=False)
def load_all_documents():

    documents = []

    if not DOCS_DIR.exists():
        return documents

    for path in sorted(
        DOCS_DIR.rglob("*")
    ):

        if not path.is_file():
            continue

        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue

        try:

            if path.suffix.lower() == ".pdf":
                sections = extract_pdf(path)
            else:
                sections = extract_docx(path)

            documents.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
                    "sections": sections,
                    "error": None,
                }
            )

        except Exception as exc:

            documents.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
                    "sections": [],
                    "error": str(exc),
                }
            )

    return documents


def normalize_text(text):

    return re.sub(
        r"\s+",
        " ",
        text or "",
    ).strip()


@st.cache_data(show_spinner=False)
def create_chunks(
    documents,
    chunk_size=1400,
    overlap=200,
):

    chunks = []

    step = max(
        1,
        chunk_size - overlap,
    )

    for document in documents:

        for section in document.get(
            "sections",
            [],
        ):

            text = normalize_text(
                section.get(
                    "text",
                    "",
                )
            )

            if not text:
                continue

            start = 0

            while start < len(text):

                chunks.append(
                    {
                        "document":
                            document["name"],

                        "page":
                            section.get("page"),

                        "text":
                            text[
                                start:
                                start + chunk_size
                            ],
                    }
                )

                start += step

    return chunks


def tokenize(text):

    return re.findall(
        r"[\w\u0600-\u06FF\u4e00-\u9fff\u0400-\u04FF]+",
        (text or "").lower(),
    )


def retrieve_chunks(
    question,
    chunks,
    top_k=7,
):

    q_tokens = set(
        tokenize(question)
    )

    if not q_tokens:
        return []

    scored = []

    for chunk in chunks:

        tokens = tokenize(
            chunk["text"]
        )

        counts = Counter(tokens)

        overlap = sum(
            counts.get(token, 0)
            for token in q_tokens
        )

        unique_overlap = len(
            q_tokens.intersection(
                counts.keys()
            )
        )

        if overlap:

            score = (
                overlap
                + unique_overlap * 2
            )

            scored.append(
                (
                    score,
                    chunk,
                )
            )

    scored.sort(
        key=lambda item: item[0],
        reverse=True,
    )

    return [
        chunk
        for _, chunk in scored[:top_k]
    ]


# ============================================================
# COMPANIES
# ============================================================

def normalize_company(company):

    if not isinstance(company, dict):
        return {}

    return {
        "name":
            str(
                company.get(
                    "name",
                    "",
                )
            ).strip(),

        "website":
            str(
                company.get(
                    "website",
                    "",
                )
            ).strip(),

        "country":
            str(
                company.get(
                    "country",
                    "",
                )
            ).strip(),

        "sector":
            str(
                company.get(
                    "sector",
                    "",
                )
            ).strip(),

        "description":
            str(
                company.get(
                    "description",
                    "",
                )
            ).strip(),

        "logo_url":
            str(
                company.get(
                    "logo_url",
                    "",
                )
            ).strip(),
    }


def load_companies():

    # --------------------------------------------------------
    # 1. LOCAL companies.json
    # --------------------------------------------------------

    if COMPANIES_FILE.exists():

        try:

            with open(
                COMPANIES_FILE,
                "r",
                encoding="utf-8",
            ) as f:

                payload = json.load(f)

            if isinstance(payload, list):

                companies = [
                    normalize_company(item)
                    for item in payload
                ]

                companies = [
                    item
                    for item in companies
                    if item.get("name")
                ]

                return companies, "local"

            if isinstance(payload, dict):

                for key in (
                    "companies",
                    "organizations",
                    "data",
                    "items",
                ):

                    value = payload.get(key)

                    if isinstance(
                        value,
                        list,
                    ):

                        companies = [
                            normalize_company(item)
                            for item in value
                        ]

                        companies = [
                            item
                            for item in companies
                            if item.get("name")
                        ]

                        return companies, "local"

        except Exception:
            pass

    # --------------------------------------------------------
    # 2. GitHub fallback
    # --------------------------------------------------------

    github_url = (
        "https://raw.githubusercontent.com/"
        "vk7496/Scocoex.-Nexus-/main/"
        "companies.json"
    )

    try:

        request = Request(
            github_url,
            headers={
                "User-Agent":
                    "SCOCOEX-NEXUS"
            },
        )

        with urlopen(
            request,
            timeout=8,
        ) as response:

            raw = response.read().decode(
                "utf-8"
            )

        payload = json.loads(raw)

        if isinstance(payload, list):

            companies = [
                normalize_company(item)
                for item in payload
            ]

        else:

            companies = []

            for key in (
                "companies",
                "organizations",
                "data",
                "items",
            ):

                if isinstance(
                    payload.get(key),
                    list,
                ):

                    companies = [
                        normalize_company(item)
                        for item in payload[key]
                    ]

                    break

        companies = [
            item
            for item in companies
            if item.get("name")
        ]

        return companies, "github"

    except Exception as exc:

        raise RuntimeError(
            f"companies.json could not be loaded: {exc}"
        )


# ============================================================
# STRATEGIC ORGANIZATIONS
# ============================================================

STRATEGIC_PROFILES = {

    "golrang": {
        "keywords": [
            "golrang",
            "گلرنگ",
        ],
        "label":
            "Golrang",
        "sector":
            "Consumer Industries • Manufacturing • Investment",
        "accent":
            "gold",
    },

    "bonyad": {
        "keywords": [
            "bonyad mostazafin",
            "bonyad-e mostazafin",
            "mostazafin",
            "foundation mostazafin",
            "بنیاد مستضعفان",
        ],
        "label":
            "Bonyad Mostazafin",
        "sector":
            "Strategic Investment • Infrastructure • Holding",
        "accent":
            "green",
    },

    "copper": {
        "keywords": [
            "copper industries",
            "national iranian copper",
            "nicico",
            "مس ایران",
            "صنایع مس ایران",
        ],
        "label":
            "National Iranian Copper Industries",
        "sector":
            "Mining • Copper • Strategic Resources",
        "accent":
            "blue",
    },
}


def strategic_key(company):

    name = (
        str(
            company.get(
                "name",
                "",
            )
        )
        .lower()
        .strip()
    )

    for key, profile in STRATEGIC_PROFILES.items():

        for keyword in profile["keywords"]:

            if keyword.lower() in name:
                return key

    return None


def strategic_company_objects(companies):

    result = []

    for key, profile in STRATEGIC_PROFILES.items():

        found = None

        for company in companies:

            if strategic_key(company) == key:

                found = company
                break

        if found is None:

            found = {
                "name":
                    profile["label"],
                "website": "",
                "country":
                    "Iran",
                "sector":
                    profile["sector"],
                "description":
                    "",
                "logo_url": "",
            }

        result.append(
            (
                key,
                found,
            )
        )

    return result


# ============================================================
# GROQ
# ============================================================

def get_groq_client():

    if Groq is None:
        return None

    api_key = None

    try:
        api_key = st.secrets.get(
            "GROQ_API_KEY"
        )
    except Exception:
        api_key = None

    if not api_key:
        api_key = os.getenv(
            "GROQ_API_KEY"
        )

    if not api_key:
        return None

    return Groq(
        api_key=api_key
    )


def ask_groq(
    question,
    context,
    mode="qa",
):

    client = get_groq_client()

    if client is None:

        return (
            None,
            t("groq_missing"),
        )

    if mode == "summary":

        system_prompt = f"""
You are SCOCOEX NEXUS, an institutional intelligence assistant.

Summarize the supplied document accurately.

Do not invent facts.

Focus on:
- organizations
- investments
- markets
- strategic information
- financial information
- dates
- opportunities
- risks

Answer in the user's selected language:
{st.session_state.language_name}

Use clear professional formatting.
"""

        user_prompt = (
            "Summarize this document:\n\n"
            + context
        )

    else:

        system_prompt = f"""
You are SCOCOEX NEXUS, an institutional AI intelligence assistant.

Your task is to answer the user's question using ONLY
the supplied SCOCOEX Knowledge Base context.

Rules:

1. Do not invent facts.
2. If the answer is not supported by the context,
   clearly say that the Knowledge Base does not contain
   sufficient information.
3. Prefer precise, concise institutional language.
4. Identify relevant organizations, sectors, countries,
   investments, dates and figures when present.
5. Answer in the user's selected language:
   {st.session_state.language_name}

The user question is below.
"""

        user_prompt = (
            "KNOWLEDGE BASE CONTEXT:\n\n"
            + context
            + "\n\n"
            + "USER QUESTION:\n"
            + question
        )

    try:

        response = client.chat.completions.create(

            model="llama-3.3-70b-versatile",

            messages=[
                {
                    "role":
                        "system",
                    "content":
                        system_prompt,
                },
                {
                    "role":
                        "user",
                    "content":
                        user_prompt,
                },
            ],

            temperature=0.15,

            max_tokens=1800,
        )

        answer = (
            response
            .choices[0]
            .message
            .content
            .strip()
        )

        return answer, None

    except Exception as exc:

        return (
            None,
            f"Groq error: {exc}",
        )


# ============================================================
# PDF OUTPUT
# ============================================================

def create_pdf(
    text,
    title="SCOCOEX AI Answer",
):

    if (
        SimpleDocTemplate is None
        or Paragraph is None
        or Spacer is None
    ):
        return None

    try:

        buffer = BytesIO()

        styles = getSampleStyleSheet()

        story = [
            Paragraph(
                html.escape(title),
                styles["Title"],
            ),
            Spacer(
                1,
                12,
            ),
        ]

        for paragraph in (
            text or ""
        ).split("\n"):

            if paragraph.strip():

                story.append(
                    Paragraph(
                        html.escape(
                            paragraph.strip()
                        ),
                        styles["BodyText"],
                    )
                )

                story.append(
                    Spacer(
                        1,
                        7,
                    )
                )

        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
        )

        doc.build(story)

        return buffer.getvalue()

    except Exception:

        return None


# ============================================================
# DATA LOAD
# ============================================================

documents = load_all_documents()

chunks = create_chunks(
    documents
)

try:

    companies, companies_source = (
        load_companies()
    )

    companies_error = None

except Exception as exc:

    companies = []

    companies_source = None

    companies_error = str(exc)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    """
    <div style="
        font-size:21px;
        font-weight:800;
        color:#FFFFFF;
        margin-bottom:4px;
    ">
        SCOCOEX NEXUS
    </div>

    <div style="
        font-size:11px;
        color:#929CB2;
        letter-spacing:.12em;
        margin-bottom:18px;
    ">
        AI KNOWLEDGE PLATFORM
    </div>
    """,
    unsafe_allow_html=True,
)


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
    key="main_navigation_v5",
    label_visibility="collapsed",
)


st.sidebar.divider()

st.sidebar.caption(
    f"{t('documents')}: {len(documents)}"
)

st.sidebar.caption(
    f"{t('international_companies')}: "
    f"{len(companies)}"
)

st.sidebar.caption(
    "SCOCOEX NEXUS • AI PLATFORM"
)


# ============================================================
# HOME
# ============================================================

if page == t("home"):

    # --------------------------------------------------------
    # LOGO
    # --------------------------------------------------------

    if LOGO_FILE.exists():

        left, center, right = st.columns(
            [1, 2, 1]
        )

        with center:

            st.image(
                str(LOGO_FILE),
                use_container_width=True,
            )

    else:

        logo_candidates = sorted(
            ASSETS_DIR.glob(
                "scocoex_logo*.png"
            )
        )

        if logo_candidates:

            left, center, right = st.columns(
                [1, 2, 1]
            )

            with center:

                st.image(
                    str(logo_candidates[0]),
                    use_container_width=True,
                )


    # --------------------------------------------------------
    # HERO
    # --------------------------------------------------------

    render_html(
        f"""
        <div
            class="hero"
            dir="{direction()}"
        >

            <div class="hero-label">
                {esc(t("hero_label"))}
            </div>

            <h1>
                {t("hero_title")}
            </h1>

            <p>
                {esc(t("hero_text"))}
            </p>

        </div>
        """
    )


    # --------------------------------------------------------
    # METRICS
    # --------------------------------------------------------

    metrics = [
        (
            t("documents"),
            len(documents),
        ),
        (
            t("chunks"),
            len(chunks),
        ),
        (
            t("international_companies"),
            len(companies),
        ),
        (
            t("languages"),
            5,
        ),
    ]


    cols = st.columns(4)

    for col, (label, value) in zip(
        cols,
        metrics,
    ):

        with col:

            render_html(
                f"""
                <div
                    class="metric-card"
                    dir="{direction()}"
                >

                    <div class="metric-label">
                        {esc(label)}
                    </div>

                    <div class="metric-number">
                        {esc(value)}
                    </div>

                </div>
                """
            )


    st.markdown("")

    # --------------------------------------------------------
    # PLATFORM FEATURES
    # --------------------------------------------------------

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            {esc(t("knowledge"))}
        </div>
        """
    )


    cards = [

        (
            "📚",
            t("knowledge"),
            t("knowledge_text"),
        ),

        (
            "🤖",
            t("ai"),
            t("ai_text"),
        ),

        (
            "🌍",
            t("network"),
            t("network_text"),
        ),
    ]


    cols = st.columns(3)

    for col, (
        icon,
        title,
        description,
    ) in zip(
        cols,
        cards,
    ):

        with col:

            render_html(
                f"""
                <div
                    class="feature-card"
                    dir="{direction()}"
                >

                    <div class="feature-icon">
                        {icon}
                    </div>

                    <div class="feature-title">
                        {esc(title)}
                    </div>

                    <div class="feature-text">
                        {esc(description)}
                    </div>

                </div>
                """
            )


# ============================================================
# INFORMATION CENTER
# ============================================================

elif page == t("info"):

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            {esc(t("information_title"))}
        </div>

        <div
            class="section-subtitle"
            dir="{direction()}"
        >
            {esc(t("information_text"))}
        </div>
        """
    )


    if not DOCS_DIR.exists():

        st.error(
            t("docs_missing")
        )

    elif not documents:

        st.warning(
            t("no_docs")
        )

    else:

        search = st.text_input(
            t("search_documents"),
            key="document_search_v5",
        )


        filtered_documents = [

            doc
            for doc in documents

            if (
                not search
                or search.lower()
                in doc["name"].lower()
            )
        ]


        for document in filtered_documents:

            with st.expander(
                f"📄 {document['name']}"
            ):

                total_words = sum(
                    len(
                        tokenize(
                            section["text"]
                        )
                    )
                    for section
                    in document["sections"]
                )


                full_text = " ".join(
                    section["text"]
                    for section
                    in document["sections"]
                )


                numeric_count = len(
                    re.findall(
                        r"\b\d+(?:[.,]\d+)?\b",
                        full_text,
                    )
                )


                c1, c2, c3 = st.columns(3)


                c1.metric(
                    t("extracted_sections"),
                    len(
                        document["sections"]
                    ),
                )


                c2.metric(
                    t("words"),
                    total_words,
                )


                c3.metric(
                    t("numeric_values"),
                    numeric_count,
                )


                if document.get("error"):

                    st.error(
                        document["error"]
                    )

                    continue


                if st.button(
                    t("summary"),
                    key=(
                        "summary_"
                        + document["name"]
                    ),
                    use_container_width=True,
                ):

                    with st.spinner(
                        t("generating")
                    ):

                        summary, error = (
                            ask_groq(
                                question="",
                                context=full_text[:24000],
                                mode="summary",
                            )
                        )


                    if error:

                        st.error(error)

                    elif summary:

                        render_html(
                            f"""
                            <div
                                class="answer-box"
                                dir="{direction()}"
                            >
                                {esc(summary).replace(chr(10), "<br>")}
                            </div>
                            """
                        )


                        pdf_data = create_pdf(
                            summary,
                            (
                                "SCOCOEX — "
                                + document["name"]
                            ),
                        )


                        if pdf_data:

                            st.download_button(
                                t("download_pdf"),
                                data=pdf_data,
                                file_name=(
                                    "scocoex_ai_summary.pdf"
                                ),
                                mime="application/pdf",
                                key=(
                                    "download_summary_"
                                    + document["name"]
                                ),
                            )


# ============================================================
# AI ASSISTANT
# ============================================================

elif page == t("assistant"):

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            {esc(t("ask_title"))}
        </div>

        <div
            class="section-subtitle"
            dir="{direction()}"
        >
            {esc(t("ask_caption"))}
        </div>
        """
    )


    question = st.text_area(
        t("your_question"),
        placeholder=(
            t("your_question")
            + "..."
        ),
        height=150,
        key="assistant_question_v5",
    )


    st.caption(
        f"{t('response_language')}: "
        f"{st.session_state.language_name}"
    )


    ask_clicked = st.button(
        t("ask_button"),
        type="primary",
        use_container_width=True,
        key="assistant_ask_button_v5",
    )


    if ask_clicked:

        if not question.strip():

            st.warning(
                t("your_question")
            )

        elif not chunks:

            st.warning(
                t("no_results")
            )

        else:

            with st.spinner(
                t("searching")
            ):

                relevant = retrieve_chunks(
                    question,
                    chunks,
                    top_k=7,
                )


            if not relevant:

                st.warning(
                    t("no_results")
                )

            else:

                context_parts = []


                for item in relevant:

                    page_info = ""

                    if item.get("page"):

                        page_info = (
                            ", page "
                            + str(
                                item["page"]
                            )
                        )


                    context_parts.append(
                        "SOURCE: "
                        + item["document"]
                        + page_info
                        + "\n"
                        + item["text"]
                    )


                context = (
                    "\n\n---\n\n"
                    .join(
                        context_parts
                    )
                )


                with st.spinner(
                    t("generating")
                ):

                    answer, error = ask_groq(
                        question,
                        context,
                        mode="qa",
                    )


                if error:

                    st.error(error)

                elif answer:

                    render_html(
                        f"""
                        <div
                            class="answer-box"
                            dir="{direction()}"
                        >
                            {esc(answer).replace(chr(10), "<br>")}
                        </div>
                        """
                    )


                    render_html(
                        f"""
                        <div
                            class="section-title"
                            style="font-size:22px;margin-top:30px;"
                            dir="{direction()}"
                        >
                            {esc(t("sources"))}
                        </div>
                        """
                    )


                    seen = set()


                    for item in relevant:

                        source_key = (
                            item["document"],
                            item.get("page"),
                        )


                        if source_key in seen:
                            continue


                        seen.add(
                            source_key
                        )


                        page_text = ""

                        if item.get("page"):

                            page_text = (
                                " — page "
                                + str(
                                    item["page"]
                                )
                            )


                        st.info(
                            "📄 "
                            + item["document"]
                            + page_text
                        )


                    pdf_data = create_pdf(
                        answer,
                        "SCOCOEX NEXUS — AI Assistant",
                    )


                    if pdf_data:

                        st.download_button(
                            t("download_pdf"),
                            data=pdf_data,
                            file_name=(
                                "scocoex_ai_answer.pdf"
                            ),
                            mime="application/pdf",
                            key="download_ai_answer_v5",
                        )


# ============================================================
# INTERNATIONAL COMPANIES
# ============================================================

elif page == t("companies"):

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            🌍 {esc(t("companies_title"))}
        </div>

        <div
            class="section-subtitle"
            dir="{direction()}"
        >
            SCOCOEX NEXUS strategic organization intelligence
            and global company network.
        </div>
        """
    )


    if companies_error:

        st.error(
            companies_error
        )


    if not companies:

        st.warning(
            t("no_companies")
        )

    else:

        # ----------------------------------------------------
        # STRATEGIC IRANIAN ORGANIZATIONS
        # ----------------------------------------------------

        strategic = strategic_company_objects(
            companies
        )


        render_html(
            f"""
            <div class="directory-header">
                <span>✦</span>
                {esc(t("strategic")).upper()}
            </div>
            """
        )


        strategic_cols = st.columns(3)


        for col, (
            strategic_id,
            company,
        ) in zip(
            strategic_cols,
            strategic,
        ):

            profile = STRATEGIC_PROFILES[
                strategic_id
            ]


            name = (
                company.get("name")
                or profile["label"]
            )


            website = (
                company.get("website")
                or ""
            )


            country = (
                company.get("country")
                or "Iran"
            )


            with col:

                link = ""

                if website:

                    link = (
                        f"""
                        <a
                            class="strategic-link"
                            href="{esc(website)}"
                            target="_blank"
                            rel="noopener noreferrer"
                        >
                            {esc(t("official"))} →
                        </a>
                        """
                    )


                render_html(
                    f"""
                    <div
                        class="strategic-card"
                        dir="{direction()}"
                    >

                        <div
                            class="strategic-badge"
                        >
                            STRATEGIC PROFILE
                        </div>

                        <div
                            class="strategic-name"
                        >
                            {esc(name)}
                        </div>

                        <div
                            class="strategic-sector"
                        >
                            {esc(profile["sector"])}
                        </div>

                        <div
                            class="strategic-copy"
                        >
                            SCOCOEX strategic profile
                            for discovery, institutional
                            relationships, investment
                            opportunities and intelligence.
                        </div>

                        {link}

                    </div>
                    """
                )


        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        render_html(
            f"""
            <div class="directory-header">
                <span>🌐</span>
                {esc(t("international_network")).upper()}
            </div>
            """
        )


        search = st.text_input(
            t("search_network"),
            key="company_search_v5",
        )


        filtered = []


        query = (
            search.strip().lower()
            if search
            else ""
        )


        for company in companies:

            searchable = " ".join(
                [
                    company.get(
                        "name",
                        "",
                    ),

                    company.get(
                        "country",
                        "",
                    ),

                    company.get(
                        "sector",
                        "",
                    ),

                    company.get(
                        "description",
                        "",
                    ),
                ]
            ).lower()


            if (
                not query
                or query in searchable
            ):

                filtered.append(
                    company
                )


        st.caption(
            f"{len(filtered)} "
            f"{t('organizations_in_network')}"
        )


        # ----------------------------------------------------
        # COMPANY GRID
        # ----------------------------------------------------

        columns = st.columns(3)


        for index, company in enumerate(
            filtered
        ):

            col = columns[
                index % 3
            ]


            name = (
                company.get("name")
                or "Organization"
            )


            country = (
                company.get("country")
                or "International"
            )


            sector = (
                company.get("sector")
                or "Strategic Organization"
            )


            description = (
                company.get("description")
                or "International organization within the SCOCOEX global intelligence network."
            )


            website = (
                company.get("website")
                or ""
            )


            initial = (
                name[:1].upper()
                if name
                else "S"
            )


            link = ""

            if website:

                link = (
                    f"""
                    <a
                        class="company-link"
                        href="{esc(website)}"
                        target="_blank"
                        rel="noopener noreferrer"
                    >
                        {esc(t("visit"))}
                    </a>
                    """
                )


            with col:

                render_html(
                    f"""
                    <div
                        class="company-card"
                        dir="{direction()}"
                    >

                        <div
                            class="company-initial"
                        >
                            {esc(initial)}
                        </div>

                        <div
                            class="company-name"
                        >
                            {esc(name)}
                        </div>

                        <div
                            class="company-sector"
                        >
                            {esc(sector)}
                        </div>

                        <div
                            class="company-country"
                        >
                            {esc(country)}
                        </div>

                        <div
                            class="company-description"
                        >
                            {esc(description)}
                        </div>

                        {link}

                    </div>
                    """
                )


# ============================================================
# INTELLIGENCE DASHBOARD
# ============================================================

elif page == t("intel"):

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            {esc(t("intel_title"))}
        </div>

        <div
            class="section-subtitle"
            dir="{direction()}"
        >
            {esc(t("intel_caption"))}
        </div>
        """
    )


    total_words = 0
    total_numbers = 0
    total_sections = 0


    for document in documents:

        for section in document.get(
            "sections",
            [],
        ):

            text = section.get(
                "text",
                "",
            )

            total_words += len(
                tokenize(text)
            )

            total_numbers += len(
                re.findall(
                    r"\b\d+(?:[.,]\d+)?\b",
                    text,
                )
            )

            total_sections += 1


    cols = st.columns(4)


    dashboard_metrics = [

        (
            t("documents"),
            len(documents),
        ),

        (
            t("chunks"),
            len(chunks),
        ),

        (
            t("words"),
            total_words,
        ),

        (
            t("numeric_values"),
            total_numbers,
        ),
    ]


    for col, (
        label,
        value,
    ) in zip(
        cols,
        dashboard_metrics,
    ):

        with col:

            render_html(
                f"""
                <div class="metric-card">

                    <div class="metric-label">
                        {esc(label)}
                    </div>

                    <div class="metric-number">
                        {esc(value)}
                    </div>

                </div>
                """
            )


    st.markdown("")


    render_html(
        f"""
        <div
            class="answer-box"
            dir="{direction()}"
        >

            <strong>
                {esc(t("documents"))}
            </strong>
            <br>

            {len(documents)}

            <br><br>

            <strong>
                {esc(t("extracted_sections"))}
            </strong>
            <br>

            {total_sections}

            <br><br>

            <strong>
                {esc(t("international_companies"))}
            </strong>
            <br>

            {len(companies)}

        </div>
        """
    )


# ============================================================
# GALLERY
# ============================================================

elif page == t("gallery"):

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            {esc(t("gallery_title"))}
        </div>

        <div
            class="section-subtitle"
            dir="{direction()}"
        >
            {esc(t("gallery_text"))}
        </div>
        """
    )


    gallery_files = []

    if ASSETS_DIR.exists():

        for file in ASSETS_DIR.iterdir():

            if file.suffix.lower() in {
                ".png",
                ".jpg",
                ".jpeg",
                ".webp",
            }:

                if (
                    file.name
                    != LOGO_FILE.name
                ):

                    gallery_files.append(
                        file
                    )


    if gallery_files:

        cols = st.columns(3)

        for index, image_file in enumerate(
            gallery_files
        ):

            with cols[index % 3]:

                st.image(
                    str(image_file),
                    use_container_width=True,
                )

    else:

        render_html(
            f"""
            <div
                class="feature-card"
                dir="{direction()}"
            >

                <div class="feature-icon">
                    🌐
                </div>

                <div class="feature-title">
                    SCOCOEX Global Week
                </div>

                <div class="feature-text">
                    {esc(t("gallery_text"))}
                </div>

            </div>
            """
        )


# ============================================================
# ABOUT
# ============================================================

elif page == t("about"):

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            {esc(t("about_title"))}
        </div>

        <div
            class="answer-box"
            dir="{direction()}"
        >

            <h3>
                SCOCOEX NEXUS
            </h3>

            <p>
                SCOCOEX NEXUS is an AI-powered
                global intelligence and knowledge
                platform designed to connect
                organizations, investors, companies
                and strategic information.
            </p>

            <p>
                The platform combines structured
                organizational intelligence with
                searchable institutional documents
                and AI-assisted knowledge discovery.
            </p>

        </div>
        """
    )


# ============================================================
# CONTACT
# ============================================================

elif page == t("contact"):

    render_html(
        f"""
        <div
            class="section-title"
            dir="{direction()}"
        >
            {esc(t("contact_title"))}
        </div>

        <div
            class="section-subtitle"
            dir="{direction()}"
        >
            Connect with SCOCOEX for strategic
            collaboration and institutional opportunities.
        </div>
        """
    )


    with st.form(
        "contact_form_v5"
    ):

        name = st.text_input(
            t("name")
        )

        organization = st.text_input(
            t("company_org")
        )

        email = st.text_input(
            t("email")
        )

        interest = st.text_input(
            t("interest")
        )

        message = st.text_area(
            t("message"),
            height=160,
        )


        submitted = st.form_submit_button(
            t("send"),
            use_container_width=True,
        )


        if submitted:

            if not name.strip():

                st.warning(
                    t("name")
                )

            elif not email.strip():

                st.warning(
                    t("email")
                )

            else:

                st.success(
                    t("received")
    )
