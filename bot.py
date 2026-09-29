import os
import csv
import random
import time
import json
import re
import html
import zipfile
import importlib
import sqlite3
import hashlib
from collections import Counter
from xml.sax.saxutils import escape
from datetime import datetime, timezone
from pathlib import Path

import requests
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import urljoin

try:
    import httpx
except ImportError:
    httpx = None

try:
    import openai
except ImportError:
    openai = None

# --- ۱. تنظیمات اولیه و بارگذاری محیط ---
# پیدا کردن مسیر دقیق فایل .env برای جلوگیری از خطای مسیر در VS Code
env_path = Path(__file__).parent / '.env'
load_dotenv(dotenv_path=env_path)

CSV_FILE = 'results.csv'
EXCEL_FILE = 'results.xlsx'
STATE_FILE = 'scraper_state.json'
MARKET_DB_FILE = 'market_data.db'
MARKET_REPORT_FILE = 'market_report.json'
MARKET_SUMMARY_CSV = 'market_summary.csv'
MARKET_DATA_CSV = 'market_data.csv'
MARKET_PUBLIC_JSON = 'public_market.json'
PROJECT_VALIDATION_DAYS = int(os.getenv('PROJECT_VALIDATION_DAYS', '30'))
MARKET_REPORT_DAYS = int(os.getenv('MARKET_REPORT_DAYS', '30'))
MARKET_TREND_WINDOW_DAYS = int(os.getenv('MARKET_TREND_WINDOW_DAYS', '7'))
MIN_DELAY_BETWEEN_SITES = float(os.getenv('MIN_DELAY_BETWEEN_SITES', '8'))
MAX_DELAY_BETWEEN_SITES = float(os.getenv('MAX_DELAY_BETWEEN_SITES', '18'))
REQUEST_TIMEOUT_DEFAULT = int(os.getenv('REQUEST_TIMEOUT_DEFAULT', '30'))
MIN_SITE_INTERVAL_SECONDS = int(os.getenv('MIN_SITE_INTERVAL_SECONDS', '600'))
CIRCUIT_FAILURE_THRESHOLD = int(os.getenv('CIRCUIT_FAILURE_THRESHOLD', '3'))
CIRCUIT_COOLDOWN_SECONDS = int(os.getenv('CIRCUIT_COOLDOWN_SECONDS', '1800'))
CACHE_MAX_AGE_SECONDS = int(os.getenv('CACHE_MAX_AGE_SECONDS', '3600'))
CACHE_DIR = Path(__file__).parent / '.scraper_cache'
CACHE_DIR.mkdir(exist_ok=True)
AI_MATCH_THRESHOLD = int(os.getenv('AI_MATCH_THRESHOLD', '80'))
AI_BATCH_SIZE = int(os.getenv('AI_BATCH_SIZE', '8'))
AI_BACKFILL_LIMIT = int(os.getenv('AI_BACKFILL_LIMIT', '20'))
USER_PROFILE = os.getenv('USER_PROFILE', '').strip()
GROQ_MODEL = os.getenv('GROQ_MODEL', 'openai/gpt-oss-20b')
RUBIKA_BOT_TOKEN = os.getenv('RUBIKA_BOT_TOKEN', '').strip()
RUBIKA_CHAT_ID = os.getenv('RUBIKA_CHAT_ID', '').strip()
RUBIKA_API_BASE = os.getenv('RUBIKA_API_BASE', 'https://botapi.rubika.ir/v3').rstrip('/')
RUBIKA_MIN_SCORE = int(os.getenv('RUBIKA_MIN_SCORE', str(AI_MATCH_THRESHOLD)))
USE_PLAYWRIGHT_FALLBACK = os.getenv('USE_PLAYWRIGHT_FALLBACK', '1').lower() in ('1', 'true', 'yes')
PLAYWRIGHT_BROWSER_CHANNELS = [x.strip() for x in os.getenv('PLAYWRIGHT_BROWSER_CHANNELS', 'msedge,chrome').split(',') if x.strip()]
DETAIL_ENRICHMENT_LIMIT = int(os.getenv('DETAIL_ENRICHMENT_LIMIT', '8'))
DETAIL_REQUEST_TIMEOUT = int(os.getenv('DETAIL_REQUEST_TIMEOUT', '20'))
DETAIL_DELAY_MIN = float(os.getenv('DETAIL_DELAY_MIN', '2'))
DETAIL_DELAY_MAX = float(os.getenv('DETAIL_DELAY_MAX', '5'))
GROQ_TOTAL_ATTEMPTS = int(os.getenv('GROQ_TOTAL_ATTEMPTS', '3'))
BOT_ROLE = os.getenv('BOT_ROLE', 'all').strip().lower()
BRAIN_API_URL = os.getenv('BRAIN_API_URL', '').strip()
BRAIN_API_TOKEN = os.getenv('BRAIN_API_TOKEN', '').strip()
BRAIN_TIMEOUT = int(os.getenv('BRAIN_TIMEOUT', '30'))
SEND_SCAN_TO_BRAIN = os.getenv('SEND_SCAN_TO_BRAIN', '0').lower() in ('1', 'true', 'yes')
BRAIN_DRY_RUN = os.getenv('BRAIN_DRY_RUN', '0').lower() in ('1', 'true', 'yes')
BRAIN_PAYLOAD_FILE = os.getenv('BRAIN_PAYLOAD_FILE', 'brain_test_payload.json')
MAX_SOURCES_PER_RUN = int(os.getenv('MAX_SOURCES_PER_RUN', '0'))
SOURCE_SCAN_MODE = os.getenv('SOURCE_SCAN_MODE', 'all').strip().lower()
ENABLE_IRANTALENT = os.getenv('ENABLE_IRANTALENT', '1').lower() in ('1', 'true', 'yes')
ENABLE_QUERA = os.getenv('ENABLE_QUERA', '1').lower() in ('1', 'true', 'yes')
ENABLE_KARBOOM = os.getenv('ENABLE_KARBOOM', '1').lower() in ('1', 'true', 'yes')

# شبکه اسکرپر: به‌صورت پیش‌فرض مستقیم است تا HTTP(S)_PROXY سیستم عامل
# بدون اطلاع کاربر وارد requests نشود و باعث خطاهای ProxyError نشود.
SCRAPER_USE_ENV_PROXY = os.getenv('SCRAPER_USE_ENV_PROXY', '0').lower() in ('1', 'true', 'yes')
SCRAPER_PROXY_URL = os.getenv('SCRAPER_PROXY_URL', '').strip()
SCRAPER_HTTP_PROXY = os.getenv('SCRAPER_HTTP_PROXY', '').strip()
SCRAPER_HTTPS_PROXY = os.getenv('SCRAPER_HTTPS_PROXY', '').strip()

# نام‌های داخلی برای سازگاری کد
CSV_HEADERS = ['Title','Company','Contact','Link','Date','Source','Match Score','Scan Date','City','Work Mode','Employment Type','Salary','Job Family','Seniority','Skills']

# سربرگ‌های فارسی برای فایل CSV و Excel
DISPLAY_HEADERS = ['عنوان شغلی','شرکت','نام مسئول / تماس','لینک','تاریخ آگهی','منبع','امتیاز سازگاری','تاریخ اسکن','شهر','نوع کار','نوع همکاری','حقوق','خانواده شغلی','سطح ارشدیت','مهارت‌ها']
HEADER_ALIASES = {
    'Title': 'Title', 'عنوان شغلی': 'Title',
    'Company': 'Company', 'شرکت': 'Company',
    'Contact': 'Contact', 'نام مسئول / تماس': 'Contact', 'مسئول استخدام': 'Contact', 'نام مسئول': 'Contact',
    'Link': 'Link', 'لینک': 'Link',
    'Date': 'Date', 'تاریخ': 'Date', 'تاریخ آگهی': 'Date',
    'Source': 'Source', 'منبع': 'Source',
    'Match Score': 'Match Score', 'امتیاز سازگاری': 'Match Score',
    'Scan Date': 'Scan Date', 'تاریخ اسکن': 'Scan Date',
}

_BS4_CLASS = None
_BS4_IMPORT_ERROR = None

# تعریف متغیر سراسری برای پروکسی‌ها (برای سازگاری با نسخه‌های قبلی)
# اسکرپر به‌صورت پیش‌فرض از این متغیر استفاده نمی‌کند؛ فقط Proxy صریح در .env معتبر است.
GLOBAL_PROXIES = None

# بررسی سلامت کلید API
GROQ_KEY = os.getenv("GROQ_API_KEY") 
if not GROQ_KEY:
    print("⚠️ خطا: کلید API در فایل .env یافت نشد!")
else:
    print("✅ کلید API با موفقیت شناسایی شد.")

GROQ_CONNECTED = False
GROQ_ACTIVE_MODEL = GROQ_MODEL

# --- ۲. زیرساخت پایداری (Retry & Responsible Requesting) ---
# User-Agent شفاف و ثابت؛ تغییر مداوم هویت مرورگر برای دورزدن ضدربات انجام نمی‌شود.
USER_AGENT = os.getenv('SCRAPER_USER_AGENT', 'JobResearchBot/1.0 (+contact via configuration)')

def get_random_headers():
    """هدر استاندارد درخواست؛ نام تابع برای سازگاری با نسخه قبلی حفظ شده است."""
    return {
        'User-Agent': USER_AGENT,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'fa-IR,fa;q=0.9,en;q=0.7',
        'Connection': 'keep-alive',
    }

HEADERS = get_random_headers()

def GAPGPTMASKTOKENytuvqodm89tX0X():
    """Session پایدار؛ Retry در لایه scrape کنترل می‌شود تا دوباره‌کاری ایجاد نشود."""
    session = requests.Session()

    # نکته مهم: requests به‌صورت پیش‌فرض Proxyهای محیط سیستم (HTTP_PROXY/HTTPS_PROXY)
    # را می‌خواند. برای اسکرپر این رفتار خاموش است مگر کاربر صریحاً فعالش کند.
    session.trust_env = SCRAPER_USE_ENV_PROXY

    # Retry داخلی urllib3 را عمداً صفر می‌کنیم؛ چون scrape_job_site خودش Retry دارد.
    # این کار از پدیده «Retry روی Retry» و پیام Max retries exceeded جلوگیری می‌کند.
    retry_strategy = Retry(
        total=0,
        connect=0,
        read=0,
        status=0,
        redirect=0,
        raise_on_status=False,
    )

    adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=10, pool_maxsize=10)
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    return session

# ایجاد یک نشست (Session) پایدار برای کل برنامه
robust_session = GAPGPTMASKTOKENytuvqodm89tX0X()

# --- ۳. پیکربندی سایت‌ها ---
# ساختار داده‌ها برای مدیریت آسان‌تر در اسکرپرها
TARGET_SITES = [
    {
        "name": "Jobinja", "enabled": True, "url": "https://jobinja.ir/jobs",
        "card_selectors": ["div.c-jobListView__item", "article.c-jobListView__item", "li.c-jobListView__item"],
        "title_selectors": ["h2.c-jobListView__title a", "h2.c-jobListView__title", "h3.c-jobListView__title a", "h3.c-jobListView__title"],
        "company_selectors": [".c-jobListView__company-name", ".c-jobListView__company", "[class*='company-name']", "[class*='company']"],
        "contact_selectors": ["[class*='recruiter']", "[class*='contact']", "[class*='employer']", "[class*='hr-name']", "[class*='responsible']"],
        "date_selectors": ["time", ".c-jobListView__updated", "[class*='date']", "[class*='updated']"],
        "job_url_patterns": [r"^/jobs(?:/|$)"], "timeout": 30, "retry": 4, "detail_enrichment_limit": 5,
    },
    {
        "name": "E-Estekhdam", "enabled": True, "url": "https://e-estekhdam.com/search",
        "card_selectors": ["article", ".job-item", ".job-card", "[class*='job-item']", "[class*='job-card']", "[class*='search-item']"],
        "title_selectors": ["h1", "h2 a", "h3 a", "a[href]"],
        "company_selectors": ["[class*='company-name']", "[class*='company']", "[class*='organization']"],
        "contact_selectors": ["[class*='recruiter']", "[class*='contact']", "[class*='employer']", "[class*='responsible']", "[class*='hr-name']"],
        "date_selectors": ["time", "[class*='date']", "[class*='published']", "[class*='created']"],
        "job_url_patterns": [r"^/[a-z0-9]{6}(?:-|$)"], "timeout": 30, "retry": 4, "detail_enrichment_limit": 4,
    },
    {
        "name": "JobVision", "enabled": True, "url": "https://jobvision.ir/jobs",
        "card_selectors": ["article", "[class*='job-card']", "[class*='job-item']", "[class*='JobCard']", "[data-testid*='job']"],
        "title_selectors": ["h1", "h2", "h3", "[class*='job-title']", "[class*='JobTitle']"],
        "company_selectors": ["[class*='company-name']", "[class*='company']", "[class*='Company']"],
        "contact_selectors": ["[class*='recruiter']", "[class*='contact']", "[class*='employer']", "[class*='responsible']", "[class*='hr-name']"],
        "date_selectors": ["time", "[class*='date']", "[class*='Date']", "[class*='published']"],
        "job_url_patterns": [r"^/jobs/\d+", r"^/jobs/job-detail/\d+", r"^/JobPost/JobPostView/\d+", r"^/JobPost/JobPostViewEn/\d+"],
        "timeout": 60, "retry": 5, "detail_enrichment_limit": 5,
    },
    {
        "name": "IranTalent", "enabled": True, "url": "https://www.irantalent.com/jobs",
        "card_selectors": ["article", "[class*='job-card']", "[class*='JobCard']", "[class*='job-item']", "[data-testid*='job']"],
        "title_selectors": ["h1", "h2", "h3", "[class*='job-title']", "[class*='JobTitle']", "a[href*='/job/']"],
        "company_selectors": ["[class*='company-name']", "[class*='company']", "[class*='Company']"],
        "contact_selectors": ["[class*='recruiter']", "[class*='contact']", "[class*='employer']"],
        "date_selectors": ["time", "[class*='date']", "[class*='published']", "[class*='created']"],
        "job_url_patterns": [r"^/job/[^/]+/\d+"], "timeout": 35, "retry": 4, "detail_enrichment_limit": 3, "env_flag": "ENABLE_IRANTALENT",
    },
    {
        "name": "Quera", "enabled": True, "url": "https://quera.org/magnet/jobs",
        "card_selectors": ["article", "[class*='job-card']", "[class*='job-item']", "[class*='JobCard']", "[data-testid*='job']"],
        "title_selectors": ["h1", "h2", "h3", "[class*='job-title']", "[class*='JobTitle']", "a[href*='/r/']", "a[href*='/careers/job/']"],
        "company_selectors": ["[class*='company-name']", "[class*='company']", "[class*='Company']"],
        "contact_selectors": ["[class*='recruiter']", "[class*='contact']", "[class*='employer']"],
        "date_selectors": ["time", "[class*='date']", "[class*='published']", "[class*='created']"],
        "job_url_patterns": [r"^/r/[A-Za-z0-9]+", r"^/careers/job/\d+"], "timeout": 35, "retry": 4, "detail_enrichment_limit": 3, "env_flag": "ENABLE_QUERA",
    },
    {
        "name": "Karboom", "enabled": True, "url": "https://karboom.io/jobs",
        "card_selectors": ["article", "[class*='job-card']", "[class*='job-item']", "[class*='JobCard']", "[data-testid*='job']"],
        "title_selectors": ["h1", "h2", "h3", "[class*='job-title']", "[class*='JobTitle']", "a[href*='/jobs/']"],
        "company_selectors": ["[class*='company-name']", "[class*='company']", "[class*='Company']"],
        "contact_selectors": ["[class*='recruiter']", "[class*='contact']", "[class*='employer']"],
        "date_selectors": ["time", "[class*='date']", "[class*='published']", "[class*='created']"],
        "job_url_patterns": [r"^/jobs/[^?#]+$"], "timeout": 35, "retry": 4, "detail_enrichment_limit": 3, "env_flag": "ENABLE_KARBOOM",
    },
]


# --- ۴. توابع کمکی ---

def set_global_proxies(proxies_dict):
    """
    یک تابع ایمن برای تغییر متغیر سراسری بدون استفاده مستقیم از global در توابع پیچیده
    """
    global GLOBAL_PROXIES
    GLOBAL_PROXIES = proxies_dict

def get_scraper_proxies():
    """پروکسی اسکرپر را فقط از تنظیمات صریح کاربر می‌گیرد."""
    if SCRAPER_PROXY_URL:
        return {
            'http': SCRAPER_PROXY_URL,
            'https': SCRAPER_PROXY_URL,
        }

    proxies = {}
    if SCRAPER_HTTP_PROXY:
        proxies['http'] = SCRAPER_HTTP_PROXY
    if SCRAPER_HTTPS_PROXY:
        proxies['https'] = SCRAPER_HTTPS_PROXY

    return proxies or None


def test_connection_with_port(port):
    """
    تست کردن پورت وارد شده توسط کاربر
    """
    proxies = {
        "http": f"http://127.0.0.1:{port}",
        "https": f"http://127.0.0.1:{port}"
    }
    try:
        # تست اتصال به گوگل با استفاده از robust_session برای بهره‌گیری از سیستم Retry شما
        robust_session.get("https://www.google.com", proxies=proxies, timeout=5)
        return proxies
    except Exception:
        return None

# --- ۵. منطق اصلی (بقیه کد شما در اینجا ادامه می‌یابد) ---
# یادآوری: اگر در توابع بعدی نیاز به تغییر GLOBAL_PROXIES داشتید، 
# از تابع set_global_proxies(new_proxies) استفاده کنید.

print("\n--- ساختار اولیه ربات با موفقیت آماده شد ---")
print("ℹ️ BeautifulSoup اکنون به‌صورت Lazy بارگذاری می‌شود تا مشکل گیرکردن هنگام شروع ربات مانع اجرا نشود.")

# --- ۶. تابع هوشمند برای یافتن هر نوع اتصال ---
def test_internet():
    """بررسی اتصال مستقیم یا پروکسی صریح؛ پروکسی‌های تصادفی localhost را خودکار انتخاب نمی‌کند."""
    global GLOBAL_PROXIES
    print("🔍 در حال بررسی اتصال اینترنت...")

    # اول اتصال مستقیم را آزمایش می‌کنیم.
    try:
        robust_session.get("https://www.google.com", timeout=5, proxies=None)
        print("✅ اتصال مستقیم برقرار است.")
        return True
    except Exception:
        print("⚠️ اتصال مستقیم برقرار نشد.")

    # اگر کاربر پروکسی را صریحاً در .env داده باشد، همان را آزمایش می‌کنیم.
    configured = get_scraper_proxies()
    if configured:
        try:
            robust_session.get("https://www.google.com", proxies=configured, timeout=5)
            GLOBAL_PROXIES = configured
            print("✅ پروکسی تنظیم‌شده قابل استفاده است.")
            return True
        except Exception as e:
            print(f"⚠️ پروکسی تنظیم‌شده قابل استفاده نیست: {e}")

    print("❌ اتصال اینترنت برای اسکرپر تأیید نشد.")
    return False

# --- ۷. تست اتصال هوش مصنوعی ---
def test_ai_connection(attempt_number=1):
    """تست واقعی مدل انتخاب‌شده؛ تلاش‌ها توسط main بین اسکن سایت‌ها پخش می‌شوند."""
    global GROQ_CONNECTED
    if not GROQ_KEY:
        return False
    client, http_client = _make_groq_client()
    if client is None:
        return False
    try:
        print(f"در حال تست اتصال به Groq (تلاش {attempt_number}/{GROQ_TOTAL_ATTEMPTS})...")
        client.chat.completions.create(
            model=GROQ_ACTIVE_MODEL,
            messages=[
                {'role': 'system', 'content': 'Reply with one word only.'},
                {'role': 'user', 'content': 'OK'},
            ],
            temperature=0,
            max_tokens=1,
        )
        GROQ_CONNECTED = True
        print(f"✅ اتصال Groq با مدل {GROQ_ACTIVE_MODEL} برقرار است.")
        return True
    except Exception as e:
        GROQ_CONNECTED = False
        message = str(e)
        if '403' in message or 'Forbidden' in message:
            print("⚠️ Groq پاسخ 403 داد؛ احتمالاً دسترسی پروژه/مدل محدود است. سه تلاش همچنان طبق برنامه انجام می‌شود، اما Retry به‌تنهایی مجوز را تغییر نمی‌دهد.")
        else:
            print(f"⚠️ خطا در اتصال به Groq: {e}")
        return False
    finally:
        if http_client is not None:
            try: http_client.close()
            except Exception: pass


# --- ۳. استخراج داده ---

def _get_beautifulsoup_class():
    """BeautifulSoup را فقط هنگام نیاز وارد می‌کند تا گیرکردن import در شروع ربات مانع اجرای برنامه نشود."""
    global _BS4_CLASS, _BS4_IMPORT_ERROR
    if _BS4_CLASS is not None or _BS4_IMPORT_ERROR is not None:
        return _BS4_CLASS
    try:
        from bs4 import BeautifulSoup as _BeautifulSoup
        _BS4_CLASS = _BeautifulSoup
        print('✅ BeautifulSoup برای پردازش HTML آماده شد.')
        return _BS4_CLASS
    except KeyboardInterrupt:
        _BS4_IMPORT_ERROR = 'KeyboardInterrupt هنگام بارگذاری BeautifulSoup'
        print('⚠️ بارگذاری BeautifulSoup قطع شد؛ Parser پشتیبان استفاده می‌شود.')
    except Exception as e:
        _BS4_IMPORT_ERROR = str(e)
        print(f'⚠️ BeautifulSoup در دسترس نیست؛ Parser پشتیبان استفاده می‌شود: {e}')
    return None


def _make_soup(content):
    bs = _get_beautifulsoup_class()
    if bs is None:
        return None
    try:
        return bs(content, 'html.parser')
    except Exception as e:
        print(f'⚠️ پردازش BeautifulSoup ناموفق بود؛ Parser پشتیبان فعال می‌شود: {e}')
        return None


def _extract_without_bs4(content, site_info):
    """Parser پشتیبان سبک برای زمانی که BeautifulSoup قابل بارگذاری نیست."""
    if isinstance(content, bytes):
        text = content.decode('utf-8', errors='ignore')
    else:
        text = str(content or '')

    jobs, seen = [], set()

    # اول JSON-LD های JobPosting را امتحان می‌کنیم؛ مستقل از CSS است.
    for raw in re.findall(r'<script[^>]+type=[\"\']application/ld\+json[\"\'][^>]*>(.*?)</script>', text, flags=re.I | re.S):
        try:
            decoded = html.unescape(raw).strip()
            obj = json.loads(decoded)
        except Exception:
            continue

        stack = [obj]
        while stack:
            item = stack.pop()
            if isinstance(item, dict):
                typ = item.get('@type')
                is_job = typ == 'JobPosting' or (isinstance(typ, list) and 'JobPosting' in typ)
                if is_job:
                    title = _clean_text(item.get('title'))
                    link = _clean_text(item.get('url'))
                    org = item.get('hiringOrganization')
                    company = _clean_text(org.get('name')) if isinstance(org, dict) else _clean_text(org)
                    date_text = _clean_text(item.get('datePosted'))
                    link = urljoin(site_info['url'], link) if link else ''
                    if title and link and _looks_like_real_job_title(title) and _matches_job_url(link, site_info) and link not in seen:
                        seen.add(link)
                        jobs.append(_new_job_dict(site_info, title, company, link, date_text))
                stack.extend(item.values())
            elif isinstance(item, list):
                stack.extend(item)

    if jobs:
        return jobs

    # سپس لینک‌های شغلی را با regexp استخراج می‌کنیم.
    anchor_pattern = re.compile(r'<a\b[^>]*href=[\"\']([^\"\']+)[\"\'][^>]*>(.*?)</a>', re.I | re.S)
    tag_pattern = re.compile(r'<[^>]+>')
    for href, inner in anchor_pattern.findall(text):
        title = _clean_text(tag_pattern.sub(' ', html.unescape(inner)))
        if not title or not _looks_like_real_job_title(title) or not _matches_job_url(href, site_info):
            continue
        link = urljoin(site_info['url'], href)
        if link in seen:
            continue
        seen.add(link)
        jobs.append(_new_job_dict(site_info, title, '', link, ''))

    return jobs


def _parse_jobs_content(content, site_info):
    soup = _make_soup(content)
    if soup is not None:
        return _extract_jobs_from_soup(soup, site_info)
    return _extract_without_bs4(content, site_info)


def _clean_text(value):
    value = html.unescape(str(value or ""))
    return re.sub(r"\s+", " ", value).strip()


def _first_text(node, selectors):
    for selector in selectors or []:
        try:
            found = node.select_one(selector)
        except Exception:
            found = None
        if found:
            text = _clean_text(found.get_text(" ", strip=True))
            if text:
                return text
    return ""


def _looks_like_real_job_title(text):
    text = _clean_text(text)
    if not text or len(text) < 3 or len(text) > 180:
        return False
    bad = {
        "جستجو", "ورود", "ثبت نام", "ثبت‌نام", "ثبت آگهی",
        "درباره ما", "تماس با ما", "رزومه ساز", "رزومه‌ساز",
        "حذف فیلترها", "مشاهده همه", "فرصت‌های شغلی",
        "فرصت های شغلی", "سوالات متداول",
    }
    return text not in bad


def _same_site_url(url, site_info):
    if not url:
        return False
    absolute = urljoin(site_info["url"], url)
    from urllib.parse import urlparse
    target = urlparse(site_info["url"]).netloc.lower()
    host = urlparse(absolute).netloc.lower()
    return host == target or host.endswith("." + target)


def _matches_job_url(url, site_info):
    if not _same_site_url(url, site_info):
        return False
    from urllib.parse import urlparse
    path = urlparse(urljoin(site_info["url"], url)).path or "/"
    return any(
        re.search(pattern, path, flags=re.I)
        for pattern in site_info.get("job_url_patterns", [])
    )


def _normalize_digits(text):
    table = str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789')
    return str(text or '').translate(table)


def _detect_city(text):
    text = _clean_text(text)
    cities = ['تهران','مشهد','اصفهان','شیراز','تبریز','کرج','قم','اهواز','رشت','کرمان','یزد','ارومیه','قزوین','ساری','بندرعباس','کرمانشاه','همدان','اردبیل','سنندج','زنجان','گرگان','بوشهر']
    for city in cities:
        if city in text:
            return city
    return ''


def _detect_work_mode(text):
    t = _clean_text(text).lower()
    if any(x in t for x in ('remote','remotely','دورکاری','کاملاً دورکاری','تماماً دورکاری')): return 'دورکاری'
    if any(x in t for x in ('hybrid','هیبرید','ترکیبی','حضوری و دورکاری')): return 'هیبرید'
    if any(x in t for x in ('onsite','on-site','حضوری','دفتر')): return 'حضوری'
    return ''


def _detect_employment_type(text):
    t = _clean_text(text).lower()
    if any(x in t for x in ('تمام وقت','تمام‌وقت','full-time','full time')): return 'تمام‌وقت'
    if any(x in t for x in ('پاره وقت','پاره‌وقت','part-time','part time')): return 'پاره‌وقت'
    if any(x in t for x in ('پروژه ای','پروژه‌ای','project','freelance','فریلنس')): return 'پروژه‌ای/فریلنس'
    if any(x in t for x in ('کارآموز','internship','intern')): return 'کارآموزی'
    return ''


def _detect_seniority(text):
    t = _clean_text(text).lower()
    if any(x in t for x in ('senior','سینیور','ارشد','lead','سرپرست')): return 'ارشد'
    if any(x in t for x in ('junior','جونیور','مبتدی','entry-level','کارآموز')): return 'مبتدی/جونیور'
    if any(x in t for x in ('mid-level','mid level','میانی','متوسط')): return 'میانی'
    if any(x in t for x in ('expert','متخصص','staff','principal')): return 'متخصص/ارشد'
    return ''


def _detect_job_family(text):
    t = _clean_text(text).lower()
    families = [
        ('هوش مصنوعی و داده',('machine learning','deep learning','data scientist','data analyst','هوش مصنوعی','یادگیری ماشین','تحلیلگر داده','داده‌کاوی','nlp')),
        ('بک‌اند',('backend','back-end','بک اند','بک‌اند','django','fastapi','node.js','spring boot','.net')),
        ('فرانت‌اند',('frontend','front-end','فرانت اند','فرانت‌اند','react','vue','angular')),
        ('DevOps و زیرساخت',('devops','sre','kubernetes','docker','terraform','cloud','زیرساخت')),
        ('موبایل',('android','ios','flutter','react native','اندروید')),
        ('امنیت',('security','cybersecurity','امنیت سایبری','penetration testing','soc')),
        ('محصول و مدیریت',('product manager','product owner','مدیر محصول','مالک محصول')),
        ('فروش و بازاریابی',('sales','marketing','فروش','بازاریابی','دیجیتال مارکتینگ')),
        ('منابع انسانی',('human resources','hr ','منابع انسانی','استخدام')),
        ('مالی و حسابداری',('accounting','finance','حسابداری','مالی')),
        ('طراحی',('ui/ux','ux','ui designer','طراح رابط','گرافیک')),
    ]
    for family, keys in families:
        if any(k in t for k in keys): return family
    return ''


def _extract_skills(text):
    t = _clean_text(text).lower()
    known = ['Python','Django','FastAPI','Flask','Java','Spring','C#','.NET','JavaScript','TypeScript','React','Vue','Angular','Node.js','PHP','Laravel','Go','Golang','Rust','SQL','PostgreSQL','MySQL','MongoDB','Redis','Docker','Kubernetes','Terraform','AWS','Azure','GCP','Git','Linux','CI/CD','DevOps','Machine Learning','Deep Learning','NLP','TensorFlow','PyTorch','Power BI','Excel','Figma','Flutter','React Native','Android','iOS','Selenium','Playwright','REST API','GraphQL']
    return ', '.join(skill for skill in known if skill.lower() in t)


def _extract_salary(text):
    t = _normalize_digits(_clean_text(text))
    if not t: return ''
    patterns = [
        r'(\d+(?:[.,]\d+)?)\s*(?:تا|-|–)\s*(\d+(?:[.,]\d+)?)\s*(?:میلیون|میلیارد)\s*(?:تومان|تومن)?',
        r'(\d+(?:[.,]\d+)?)\s*(?:میلیون)\s*(?:تومان|تومن)?',
        r'(\d{1,3}(?:[,.]\d{3})+)\s*(?:تومان|تومن|ریال)',
    ]
    for pattern in patterns:
        match = re.search(pattern, t, flags=re.I)
        if match: return _clean_text(match.group(0))
    return ''


def _parse_salary_range(salary_text):
    """استخراج بازه حقوق فقط وقتی واحد از متن آگهی قابل تشخیص باشد."""
    text = _normalize_digits(_clean_text(salary_text)).replace(',', '').replace('،', '')
    if not text:
        return None, None, ''
    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:تا|-|–)\s*(\d+(?:\.\d+)?)\s*میلیون\s*(?:تومان|تومن)?', text, re.I)
    if m:
        return float(m.group(1)), float(m.group(2)), 'میلیون تومان'
    m = re.search(r'(\d+(?:\.\d+)?)\s*میلیون\s*(?:تومان|تومن)?', text, re.I)
    if m:
        value = float(m.group(1))
        return value, value, 'میلیون تومان'
    m = re.search(r'(\d{4,})\s*(?:تا|-|–)\s*(\d{4,})\s*(تومان|تومن|ریال)', text, re.I)
    if m:
        low, high = float(m.group(1)), float(m.group(2))
        if m.group(3) == 'ریال':
            low /= 10
            high /= 10
        return low / 1_000_000, high / 1_000_000, 'میلیون تومان'
    m = re.search(r'(\d{4,})\s*(تومان|تومن|ریال)', text, re.I)
    if m:
        value = float(m.group(1))
        if m.group(2) == 'ریال':
            value /= 10
        return value / 1_000_000, value / 1_000_000, 'میلیون تومان'
    return None, None, ''


def _median(values):
    values = sorted(float(v) for v in values if v is not None)
    if not values:
        return None
    mid = len(values) // 2
    if len(values) % 2:
        return values[mid]
    return (values[mid - 1] + values[mid]) / 2

def _apply_local_metadata(job, extra_text=''):
    combined = _clean_text(' '.join(str(job.get(k,'')) for k in ('Title','Company','Description','Salary','Work Mode','Employment Type','Skills')) + ' ' + extra_text)
    work_mode = _clean_text(job.get('Work Mode'))
    if work_mode:
        upper_mode = work_mode.upper()
        if upper_mode == 'TELECOMMUTE' or 'REMOTE' in upper_mode:
            work_mode = 'دورکاری'
        elif 'HYBRID' in upper_mode:
            work_mode = 'هیبرید'
        elif 'ONSITE' in upper_mode:
            work_mode = 'حضوری'
    employment_type = _clean_text(job.get('Employment Type'))
    if employment_type.upper() == 'FULL-TIME':
        employment_type = 'تمام‌وقت'
    elif employment_type.upper() == 'PART-TIME':
        employment_type = 'پاره‌وقت'
    elif employment_type.upper() in ('CONTRACTOR','CONTRACT'):
        employment_type = 'قراردادی'
    job['City'] = job.get('City') or _detect_city(combined)
    job['Work Mode'] = work_mode or _detect_work_mode(combined)
    job['Employment Type'] = employment_type or _detect_employment_type(combined)
    job['Seniority'] = job.get('Seniority') or _detect_seniority(combined)
    job['Job Family'] = job.get('Job Family') or _detect_job_family(combined)
    job['Skills'] = job.get('Skills') or _extract_skills(combined)
    job['Salary'] = job.get('Salary') or _extract_salary(combined)
    return job


def _new_job_dict(site_info, title, company='', link='', date_text='', contact='', extra_text=''):
    job = {
        'Title': _clean_text(title), 'Company': _clean_text(company), 'Contact': _clean_text(contact),
        'Link': urljoin(site_info['url'], link) if link else '', 'Date': _clean_text(date_text), 'Source': site_info['name'],
        'Match Score': '', 'Scan Date': datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds'),
        'City': '', 'Work Mode': '', 'Employment Type': '', 'Salary': '', 'Job Family': '', 'Seniority': '', 'Skills': '',
        'Description': _clean_text(extra_text)[:6000],
    }
    return _apply_local_metadata(job, extra_text)


def _contact_from_jsonld(item):
    contact_point = item.get("contactPoint")
    if isinstance(contact_point, dict):
        return _clean_text(contact_point.get("name") or contact_point.get("email"))
    if isinstance(contact_point, list):
        for point in contact_point:
            if isinstance(point, dict):
                value = _clean_text(point.get("name") or point.get("email"))
                if value:
                    return value
    return _clean_text(item.get("recruiter") or item.get("hiringManager"))


def _structured_jobposting_fields(item):
    data = {'Company':'','Contact':'','Date':'','City':'','Work Mode':'','Employment Type':'','Salary':'','Skills':'','Description':''}
    org = item.get('hiringOrganization')
    data['Company'] = _clean_text(org.get('name')) if isinstance(org,dict) else _clean_text(org)
    data['Contact'] = _contact_from_jsonld(item)
    data['Date'] = _clean_text(item.get('datePosted'))
    data['Description'] = _clean_text(item.get('description'))[:6000]
    location = item.get('jobLocation')
    if isinstance(location,list): location = location[0] if location else None
    if isinstance(location,dict):
        address = location.get('address') or {}
        if isinstance(address,dict): data['City'] = _clean_text(address.get('addressLocality'))
        elif isinstance(address,str): data['City'] = _clean_text(address)
    data['Employment Type'] = _clean_text(item.get('employmentType'))
    data['Work Mode'] = _clean_text(item.get('jobLocationType'))
    salary = item.get('baseSalary') or item.get('estimatedSalary')
    if isinstance(salary,dict):
        value = salary.get('value') or salary
        if isinstance(value,dict):
            low=value.get('minValue'); high=value.get('maxValue'); currency=salary.get('currency') or value.get('currency') or ''
            if low is not None and high is not None: data['Salary']=f'{low} - {high} {currency}'.strip()
            elif low is not None or high is not None: data['Salary']=f'{low if low is not None else high} {currency}'.strip()
        elif value is not None: data['Salary']=_clean_text(value)
    elif isinstance(salary,list) and salary: data['Salary']=_clean_text(salary[0])
    data['Skills'] = _clean_text(item.get('skills') or '')
    return data


def _extract_jsonld_jobs(soup, site_info):
    raw_jobs=[]
    def walk(obj):
        if isinstance(obj,dict):
            typ=obj.get('@type')
            if typ=='JobPosting' or (isinstance(typ,list) and 'JobPosting' in typ): raw_jobs.append(obj)
            for value in obj.values(): walk(value)
        elif isinstance(obj,list):
            for value in obj: walk(value)
    for script in soup.find_all('script',type='application/ld+json'):
        raw=(script.string or script.get_text()).strip()
        if not raw: continue
        try: walk(json.loads(raw))
        except Exception: continue
    jobs=[]; seen=set()
    for item in raw_jobs:
        title=_clean_text(item.get('title')); link=_clean_text(item.get('url')); fields=_structured_jobposting_fields(item)
        link=urljoin(site_info['url'],link) if link else ''
        if not (_looks_like_real_job_title(title) and link and _matches_job_url(link,site_info) and link not in seen): continue
        seen.add(link)
        job=_new_job_dict(site_info,title,fields['Company'],link,fields['Date'],fields['Contact'],fields['Description'])
        for key in ('City','Work Mode','Employment Type','Salary','Skills'):
            if fields.get(key): job[key]=_clean_text(fields[key])
        jobs.append(_apply_local_metadata(job,fields['Description']))
    return jobs


def _extract_anchor_jobs(soup, site_info):
    jobs=[]; seen=set()
    for anchor in soup.find_all('a',href=True):
        href=anchor.get('href','').strip(); title=_clean_text(anchor.get_text(' ',strip=True))
        if not href or not _matches_job_url(href,site_info) or not _looks_like_real_job_title(title): continue
        link=urljoin(site_info['url'],href)
        if link in seen: continue
        seen.add(link); container=anchor
        for _ in range(5):
            parent=getattr(container,'parent',None)
            if parent is None: break
            container=parent
            if getattr(parent,'name','') in ('article','li','section'): break
        text=container.get_text(' ',strip=True)[:6000]
        company=_first_text(container,site_info.get('company_selectors'))
        date_text=_first_text(container,site_info.get('date_selectors'))
        contact=_first_text(container,site_info.get('contact_selectors'))
        jobs.append(_new_job_dict(site_info,title,company,link,date_text,contact,text))
    return jobs


def _extract_card_jobs(soup, site_info):
    cards=[]
    for selector in site_info.get('card_selectors',[]):
        try: cards=soup.select(selector)
        except Exception: cards=[]
        if cards: break
    if not cards:
        for selector in site_info.get('title_selectors',[]):
            try: nodes=soup.select(selector)
            except Exception: nodes=[]
            for node in nodes:
                parent=node
                for _ in range(5):
                    parent=getattr(parent,'parent',None)
                    if parent is None: break
                    if getattr(parent,'name','') in ('article','li','section'): break
                if parent is not None: cards.append(parent)
            if cards: break
    jobs=[]; seen=set()
    for card in cards:
        title=_first_text(card,site_info.get('title_selectors'))
        if not _looks_like_real_job_title(title): continue
        link_tag=card.find('a',href=True); link=urljoin(site_info['url'],link_tag.get('href','')) if link_tag else ''
        if not link or not _matches_job_url(link,site_info) or link in seen: continue
        seen.add(link)
        company=_first_text(card,site_info.get('company_selectors')); date_text=_first_text(card,site_info.get('date_selectors')); contact=_first_text(card,site_info.get('contact_selectors'))
        time_tag=card.find('time')
        if time_tag: date_text=time_tag.get('datetime') or time_tag.get_text(' ',strip=True) or date_text
        text=card.get_text(' ',strip=True)[:6000]
        jobs.append(_new_job_dict(site_info,title,company,link,date_text,contact,text))
    return jobs


def _extract_jobs_from_soup(soup, site_info):
    jobs = _extract_card_jobs(soup, site_info)
    if jobs:
        return jobs
    jobs = _extract_jsonld_jobs(soup, site_info)
    if jobs:
        return jobs
    return _extract_anchor_jobs(soup, site_info)


def _render_page_with_playwright(url, timeout_ms=60000):
    if not USE_PLAYWRIGHT_FALLBACK: return None
    try:
        playwright_api=importlib.import_module('playwright.sync_api')
        sync_playwright=getattr(playwright_api,'sync_playwright')
    except Exception:
        print('ℹ️ Playwright نصب نیست؛ fallback مرورگر انجام نشد.')
        return None
    last_error=None
    try:
        with sync_playwright() as p:
            browser=None; chosen=None
            for channel in PLAYWRIGHT_BROWSER_CHANNELS:
                try:
                    browser=p.chromium.launch(channel=channel,headless=True); chosen=channel; break
                except Exception as e: last_error=e
            if browser is None:
                try:
                    browser=p.chromium.launch(headless=True); chosen='bundled-chromium'
                except Exception as e: last_error=e; raise
            print(f'🖥️ Playwright مرورگر {chosen} را برای {url} اجرا کرد.')
            page=browser.new_page(user_agent=USER_AGENT,locale='fa-IR',viewport={'width':1440,'height':1000})
            page.goto(url,wait_until='domcontentloaded',timeout=timeout_ms)
            try: page.wait_for_load_state('networkidle',timeout=min(timeout_ms,15000))
            except Exception: pass
            page.wait_for_timeout(2000)
            content=page.content().encode('utf-8',errors='ignore')
            browser.close(); return content
    except Exception as e:
        print(f'⚠️ رندر مرورگر برای {url} ناموفق بود: {e}')
        if last_error: print(f'ℹ️ آخرین خطای browser channel: {last_error}')
        return None


def _extract_jsonld_detail(soup):
    for script in soup.find_all('script',type='application/ld+json'):
        raw=(script.string or script.get_text()).strip()
        if not raw: continue
        try: obj=json.loads(raw)
        except Exception: continue
        stack=[obj]
        while stack:
            item=stack.pop()
            if isinstance(item,dict):
                typ=item.get('@type')
                if typ=='JobPosting' or (isinstance(typ,list) and 'JobPosting' in typ):
                    data=_structured_jobposting_fields(item); data['Title']=_clean_text(item.get('title')); return data
                stack.extend(item.values())
            elif isinstance(item,list): stack.extend(item)
    return {}


def _extract_detail_fields(soup, site_info):
    data=_extract_jsonld_detail(soup)
    if not data: data={'Title':'','Company':'','Contact':'','Date':'','City':'','Work Mode':'','Employment Type':'','Salary':'','Skills':'','Description':''}
    for key, skey in (('Title','title_selectors'),('Company','company_selectors'),('Contact','contact_selectors'),('Date','date_selectors')):
        if not data.get(key): data[key]=_first_text(soup,site_info.get(skey))
    full_text=soup.get_text(' ',strip=True); meta_map={}
    for meta in soup.find_all('meta'):
        key=_clean_text(meta.get('itemprop') or meta.get('name') or meta.get('property')); value=_clean_text(meta.get('content'))
        if key and value: meta_map[key.lower()]=value
    data['Date']=data.get('Date') or meta_map.get('dateposted') or meta_map.get('article:published_time') or meta_map.get('date') or ''
    data['Description']=data.get('Description') or meta_map.get('description') or meta_map.get('og:description') or ''
    data['City']=data.get('City') or _detect_city(full_text); data['Work Mode']=data.get('Work Mode') or _detect_work_mode(full_text)
    data['Employment Type']=data.get('Employment Type') or _detect_employment_type(full_text); data['Salary']=data.get('Salary') or _extract_salary(full_text)
    data['Skills']=data.get('Skills') or _extract_skills(full_text)
    return {k:_clean_text(v) for k,v in data.items()}


def _enrich_job_from_detail(job, site_info):
    if not job.get("Link"):
        return job
    missing = [key for key in ('Company','Contact','Date','City','Work Mode','Employment Type','Salary','Skills') if not job.get(key)]
    if not missing:
        return job
    try:
        response = robust_session.get(
            job["Link"], headers=get_random_headers(),
            proxies=get_scraper_proxies(), timeout=DETAIL_REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        soup = _make_soup(response.content)
        if soup is None:
            return job
        fields = _extract_detail_fields(soup, site_info)
        for key in missing + ['Description']:
            if fields.get(key): job[key] = fields[key]
        _apply_local_metadata(job)
    except Exception as e:
        print(f"ℹ️ جزئیات {site_info['name']} برای یک آگهی قابل دریافت نبود: {e}")
    return job


def _enrich_jobs_from_details(jobs, site_info):
    site_limit = int(site_info.get('detail_enrichment_limit', DETAIL_ENRICHMENT_LIMIT))
    if not jobs or site_limit <= 0:
        return jobs
    checked = 0
    for job in jobs:
        if checked >= site_limit:
            break
        if not any(not job.get(key) for key in ('Company','Contact','Date','City','Work Mode','Employment Type','Salary','Skills')):
            continue
        if checked > 0:
            time.sleep(random.uniform(DETAIL_DELAY_MIN, DETAIL_DELAY_MAX))
        _enrich_job_from_detail(job, site_info)
        checked += 1
    if checked:
        print(f"🧾 {site_info['name']}: برای {checked} آگهی جزئیات تکمیلی بررسی شد.")
    return jobs


def _site_cache_path(site_name):
    safe_name = ''.join(ch if ch.isalnum() else '_' for ch in site_name)
    return CACHE_DIR / f"{safe_name}.html"


def _load_cached_html(site_name):
    path = _site_cache_path(site_name)
    try:
        if not path.exists():
            return None
        age = time.time() - path.stat().st_mtime
        if age > CACHE_MAX_AGE_SECONDS:
            return None
        return path.read_bytes()
    except OSError:
        return None


def _save_cached_html(site_name, content):
    try:
        _site_cache_path(site_name).write_bytes(content)
    except OSError as e:
        print(f"⚠️ ذخیره Cache برای {site_name} ناموفق بود: {e}")


def _site_can_be_scanned(site_info, state):
    name = site_info['name']
    now = time.time()
    info = state.setdefault('site_health', {}).setdefault(name, {})

    last_scan = info.get('last_attempt_epoch', 0)
    if last_scan and now - float(last_scan) < MIN_SITE_INTERVAL_SECONDS:
        remaining = MIN_SITE_INTERVAL_SECONDS - (now - float(last_scan))
        print(f"⏭️ {name}: هنوز فاصله حداقل اسکن نگذشته؛ {remaining:.0f} ثانیه باقی مانده.")
        return False

    opened_until = float(info.get('circuit_open_until_epoch', 0) or 0)
    if opened_until > now:
        remaining = opened_until - now
        print(f"⏸️ {name}: Circuit Breaker فعال است؛ {remaining:.0f} ثانیه دیگر.")
        return False

    return True


def _record_site_success(state, site_name):
    info = state.setdefault('site_health', {}).setdefault(site_name, {})
    info['failure_count'] = 0
    info['last_success_epoch'] = time.time()
    info.pop('circuit_open_until_epoch', None)


def _record_site_failure(state, site_name):
    info = state.setdefault('site_health', {}).setdefault(site_name, {})
    failures = int(info.get('failure_count', 0)) + 1
    info['failure_count'] = failures

    if failures >= CIRCUIT_FAILURE_THRESHOLD:
        info['circuit_open_until_epoch'] = time.time() + CIRCUIT_COOLDOWN_SECONDS
        print(
            f"🛑 {site_name}: پس از {failures} خطای متوالی، "
            f"دسترسی موقتاً تا {CIRCUIT_COOLDOWN_SECONDS} ثانیه متوقف شد."
        )


def _retry_after_seconds(response):
    value = response.headers.get('Retry-After')
    if not value:
        return None
    try:
        return max(0, min(float(value), 3600))
    except ValueError:
        return None


def scrape_job_site(site_info, state=None):
    if state is None:
        state = _load_state()
    init_market_database()
    try:
        started=datetime.fromisoformat(state.get('project_started_at',datetime.now(timezone.utc).isoformat()))
        days_running=max(0,(datetime.now(timezone.utc)-started).days)
    except Exception:
        days_running=0
    print(f'📅 فاز اعتبارسنجی داده: روز {min(days_running+1,PROJECT_VALIDATION_DAYS)} از {PROJECT_VALIDATION_DAYS}')

    name = site_info['name']
    if not _site_can_be_scanned(site_info, state):
        return []

    health = state.setdefault('site_health', {}).setdefault(name, {})
    health['last_attempt_epoch'] = time.time()
    _save_state(state)

    extracted_data = []
    current_proxies = get_scraper_proxies()
    max_attempts = max(1, int(site_info.get('retry', 3)))
    last_error = None

    for attempt in range(1, max_attempts + 1):
        try:
            # Jitter محدود و کنترل‌شده؛ هدف کاهش burst است، نه مخفی‌کردن هویت ربات.
            if attempt > 1:
                delay = min(30, (2 ** (attempt - 1)) + random.uniform(1, 4))
                print(f"⏳ تلاش مجدد {name} در {delay:.1f} ثانیه...")
                time.sleep(delay)

            response = robust_session.get(
                site_info['url'],
                headers=get_random_headers(),
                proxies=current_proxies,
                timeout=site_info.get('timeout', REQUEST_TIMEOUT_DEFAULT)
            )

            retry_after = _retry_after_seconds(response)
            if response.status_code == 429:
                if retry_after is not None:
                    print(f"⏳ {name}: سایت Retry-After={retry_after:.0f}s اعلام کرده است.")
                    # برای زمان‌های کوتاه منتظر می‌مانیم؛ برای زمان‌های طولانی،
                    # درخواست‌های بیشتر را متوقف می‌کنیم تا سایت را تحت فشار نگذاریم.
                    if retry_after > 60:
                        last_error = requests.HTTPError(
                            f"HTTP 429 Too Many Requests؛ Retry-After={retry_after:.0f}s",
                            response=response
                        )
                        break
                    time.sleep(retry_after)
                raise requests.HTTPError(
                    "HTTP 429 Too Many Requests",
                    response=response
                )

            if response.status_code in (403, 451):
                raise requests.HTTPError(
                    f"HTTP {response.status_code}: دسترسی توسط سایت محدود شده است.",
                    response=response
                )

            response.raise_for_status()

            # Cache آخرین پاسخ موفق برای fallback در خطاهای بعدی.
            _save_cached_html(name, response.content)
            _record_site_success(state, name)
            _save_state(state)

            soup_jobs = _parse_jobs_content(response.content, site_info)
            extracted_data = soup_jobs

            if not extracted_data and USE_PLAYWRIGHT_FALLBACK:
                print(f"🧩 {name}: HTML معمولی آگهی نداد؛ fallback مرورگر در حال بررسی است...")
                rendered = _render_page_with_playwright(
                    site_info['url'],
                    timeout_ms=max(30000, int(site_info.get('timeout', 30) * 1000))
                )
                if rendered:
                    extracted_data = _parse_jobs_content(rendered, site_info)

            if extracted_data:
                extracted_data = _enrich_jobs_from_details(extracted_data, site_info)
                print(f"📌 {name}: {len(extracted_data)} مورد از صفحه استخراج شد.")
            else:
                print(f"⚠️ {name}: پاسخ دریافت شد، اما آگهی قابل استخراجی پیدا نشد.")

            return extracted_data

        except requests.Timeout as e:
            last_error = e
            print(f"⚠️ Timeout در {name} (تلاش {attempt}/{max_attempts}).")
        except requests.HTTPError as e:
            last_error = e
            status = getattr(getattr(e, 'response', None), 'status_code', None)

            # 403/451 را با Retry بیشتر بمباران نمی‌کنیم.
            if status in (403, 451):
                print(f"⚠️ {name}: سایت دسترسی را محدود کرده است ({status}).")
                break

            print(f"⚠️ HTTP error در {name} (تلاش {attempt}/{max_attempts}): {e}")
        except requests.RequestException as e:
            last_error = e
            print(f"⚠️ خطای درخواست در {name} (تلاش {attempt}/{max_attempts}): {e}")
        except Exception as e:
            last_error = e
            print(f"⚠️ خطای پردازش {name} (تلاش {attempt}/{max_attempts}): {e}")
            break

    _record_site_failure(state, name)
    _save_state(state)

    # اگر سایت موقتاً در دسترس نبود، آخرین HTML سالم را فقط برای همین چرخه
    # به‌عنوان fallback استفاده می‌کنیم؛ این باعث افزایش درخواست به سایت نمی‌شود.
    cached = _load_cached_html(name)
    if cached:
        print(f"♻️ استفاده از آخرین Cache سالم برای {name}.")
        try:
            return _parse_jobs_content(cached, site_info)
        except Exception as e:
            print(f"⚠️ Cache سایت {name} قابل پردازش نیست: {e}")

    print(f"⚠️ اسکن {name} پس از Retryها ناموفق بود: {last_error}")
    return []


# --- ۴. ذخیره داده‌ها ---
def load_existing_jobs():
    """خواندن results.csv با پشتیبانی از سربرگ قدیمی انگلیسی و سربرگ جدید فارسی."""
    all_data = []
    if not os.path.exists(CSV_FILE):
        return all_data
    try:
        with open(CSV_FILE, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if not header:
                return all_data

            normalized = [HEADER_ALIASES.get(_clean_text(h), '') for h in header]
            index_map = {name: idx for idx, name in enumerate(normalized) if name in CSV_HEADERS}

            for raw_row in reader:
                row = {}
                for key in CSV_HEADERS:
                    idx = index_map.get(key)
                    value = raw_row[idx] if idx is not None and idx < len(raw_row) else ''
                    row[key] = _clean_text(value)
                if row.get('Link') and row.get('Title'):
                    all_data.append(row)
    except (OSError, csv.Error) as e:
        print(f'⚠️ خطا در خواندن {CSV_FILE}: {e}')
    return all_data


def _xlsx_col_letter(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _xlsx_inline_string(cell_ref, value, style_id=0):
    text = escape(str(value if value is not None else ""))
    return (
        f'<c r="{cell_ref}" s="{style_id}" t="inlineStr">'
        f'<is><t xml:space="preserve">{text}</t></is></c>'
    )


def _xlsx_number(cell_ref, value, style_id=0):
    return f'<c r="{cell_ref}" s="{style_id}"><v>{value}</v></c>'


def export_results_to_xlsx(rows):
    """ساخت Excel خوانا بدون کتابخانه جانبی."""
    headers = DISPLAY_HEADERS
    data_rows = [
        [row.get(CSV_HEADERS[index], "") for index, _ in enumerate(headers)]
        for row in rows
        if isinstance(row, dict)
    ]

    last_row = max(1, len(data_rows) + 1)
    xml_rows = []

    header_cells = [
        _xlsx_inline_string(
            f"{_xlsx_col_letter(index)}1",
            header,
            style_id=1
        )
        for index, header in enumerate(headers, 1)
    ]
    xml_rows.append(
        '<row r="1" ht="28" customHeight="1">'
        + "".join(header_cells)
        + "</row>"
    )

    for row_index, values in enumerate(data_rows, 2):
        cells = []

        for col_index, value in enumerate(values, 1):
            ref = f"{_xlsx_col_letter(col_index)}{row_index}"

            if col_index == 7:
                try:
                    score = int(str(value).strip())
                    style = 3 if score < 60 else (4 if score < 80 else 5)
                    cells.append(_xlsx_number(ref, score, style))
                    continue
                except Exception:
                    value = ""

            style = 6 if col_index == 4 else 2
            cells.append(_xlsx_inline_string(ref, value, style))

        xml_rows.append(
            f'<row r="{row_index}" ht="42" customHeight="1">'
            + "".join(cells)
            + "</row>"
        )

    sheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:O__LAST_ROW__"/>
  <sheetViews>
    <sheetView workbookViewId="0" rightToLeft="1" showGridLines="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="34" customWidth="1"/>
    <col min="2" max="2" width="25" customWidth="1"/>
    <col min="3" max="3" width="25" customWidth="1"/>
    <col min="4" max="4" width="58" customWidth="1"/>
    <col min="5" max="5" width="20" customWidth="1"/>
    <col min="6" max="6" width="18" customWidth="1"/>
    <col min="7" max="7" width="17" customWidth="1"/>
    <col min="8" max="8" width="24" customWidth="1"/>
    <col min="9" max="9" width="16" customWidth="1"/>
    <col min="10" max="10" width="16" customWidth="1"/>
    <col min="11" max="11" width="18" customWidth="1"/>
    <col min="12" max="12" width="20" customWidth="1"/>
    <col min="13" max="13" width="22" customWidth="1"/>
    <col min="14" max="14" width="18" customWidth="1"/>
    <col min="15" max="15" width="42" customWidth="1"/>
  </cols>
  <sheetData>
    __ROWS__
  </sheetData>
  <autoFilter ref="A1:O__LAST_ROW__"/>
</worksheet>"""
    sheet_xml = sheet_xml.replace("__LAST_ROW__", str(last_row))
    sheet_xml = sheet_xml.replace("__ROWS__", "".join(xml_rows))

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
    <font><color rgb="FF0066CC"/><u/><sz val="11"/><name val="Calibri"/></font>
    <font><color rgb="FF444444"/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9E2F3"/></left>
      <right style="thin"><color rgb="FFD9E2F3"/></right>
      <top style="thin"><color rgb="FFD9E2F3"/></top>
      <bottom style="thin"><color rgb="FFD9E2F3"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  </cellXfs>
</styleSheet>"""

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="فرصت‌های شغلی" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""

    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

    try:
        with zipfile.ZipFile(EXCEL_FILE, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", content_types)
            archive.writestr("_rels/.rels", root_rels)
            archive.writestr("xl/workbook.xml", workbook_xml)
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
            archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
            archive.writestr("xl/styles.xml", styles_xml)
        print(f"📊 فایل Excel خوانا ساخته شد: {EXCEL_FILE}")
    except Exception as e:
        print(f"⚠️ ساخت فایل Excel ناموفق بود: {e}")


def save_and_sort_data(new_data):
    """ذخیره نتایج با سربرگ فارسی و ادغام فیلدهای غیرخالی با رکوردهای قبلی."""
    existing = load_existing_jobs()
    merged = {}
    for row in existing:
        link = row.get('Link')
        if link:
            merged[link] = {key: str(row.get(key, '') or '') for key in CSV_HEADERS}
    for row in new_data:
        if not isinstance(row, dict) or not row.get('Link') or not row.get('Title'):
            continue
        link = row['Link']
        merged.setdefault(link, {key: '' for key in CSV_HEADERS})
        for key in CSV_HEADERS:
            value = str(row.get(key, '') or '').strip()
            if value:
                merged[link][key] = value
    def _sort_key(row):
        try:
            score = int(str(row.get('Match Score') or 0))
        except ValueError:
            score = -1
        return (-score, row.get('Source', '').casefold(), row.get('Title', '').casefold())
    sorted_data = sorted(merged.values(), key=_sort_key)
    try:
        with open(CSV_FILE, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(DISPLAY_HEADERS)
            for row in sorted_data:
                writer.writerow([row.get(key, '') for key in CSV_HEADERS])
        print(f'💾 فایل {CSV_FILE} با {len(sorted_data)} مورد به‌روز شد؛ سربرگ‌ها فارسی هستند.')
        export_results_to_xlsx(sorted_data)
        record_market_observations(new_data)
        generate_market_report()
    except OSError as e:
        print(f'❌ خطا در ذخیره‌سازی {CSV_FILE}: {e}')


def _make_groq_client():
    if not GROQ_KEY or openai is None:
        return None, None
    http_client = None
    try:
        if httpx is not None:
            if GLOBAL_PROXIES:
                proxy_url = GLOBAL_PROXIES.get('https') or GLOBAL_PROXIES.get('http')
                http_client = httpx.Client(proxy=proxy_url, timeout=45.0, trust_env=False)
            else:
                # Groq می‌تواند از Proxyهای محیط سیستم استفاده کند؛ این بخش از اسکرپر جداست.
                http_client = httpx.Client(timeout=45.0, trust_env=True)
        kwargs = {'api_key': GROQ_KEY, 'base_url': 'https://api.groq.com/openai/v1'}
        if http_client is not None:
            kwargs['http_client'] = http_client
        return openai.OpenAI(**kwargs), http_client
    except Exception as e:
        print(f"⚠️ ساخت کلاینت Groq ناموفق بود: {e}")
        if http_client:
            http_client.close()
        return None, None


def analyze_jobs_with_ai(jobs):
    """امتیازدهی کاربر و استانداردسازی متادیتای بازار با Groq."""
    global GROQ_CONNECTED
    if not jobs or not GROQ_KEY or openai is None: return jobs
    client,http_client=_make_groq_client()
    if client is None: return jobs
    try:
        for start_idx in range(0,len(jobs),AI_BATCH_SIZE):
            batch=jobs[start_idx:start_idx+AI_BATCH_SIZE]
            payload=[{'id':i,'title':j.get('Title',''),'company':j.get('Company',''),'description':j.get('Description','')[:3500],'city':j.get('City',''),'work_mode':j.get('Work Mode',''),'employment_type':j.get('Employment Type',''),'salary':j.get('Salary',''),'source':j.get('Source','')} for i,j in enumerate(batch)]
            profile=USER_PROFILE if USER_PROFILE else 'پروفایل کاربر ارائه نشده است؛ Match Score را خالی بگذار.'
            prompt=f"""پروفایل کاربر:
{profile}

فرصت‌های شغلی:
{json.dumps(payload,ensure_ascii=False)}

برای هر شغل JSON بده: score (۰ تا ۱۰۰ فقط با وجود پروفایل)، skills، job_family، seniority، city، work_mode، employment_type و salary. هیچ داده‌ای را حدس نزن. فقط JSON معتبر با ساختار {{"results":[{{"id":0,"score":0,"skills":"","job_family":"","seniority":"","city":"","work_mode":"","employment_type":"","salary":""}}]}} برگردان."""
            response=client.chat.completions.create(model=GROQ_ACTIVE_MODEL,messages=[{'role':'system','content':'You are a labor-market data normalization and job matching assistant. Return only valid JSON.'},{'role':'user','content':prompt}],temperature=0.1,response_format={'type':'json_object'})
            result=json.loads(response.choices[0].message.content or '{}')
            mapping={'skills':'Skills','job_family':'Job Family','seniority':'Seniority','city':'City','work_mode':'Work Mode','employment_type':'Employment Type','salary':'Salary'}
            for item in result.get('results',[]):
                try: idx=int(item.get('id',-1))
                except Exception: idx=-1
                if not (0<=idx<len(batch)): continue
                target=batch[idx]
                if USER_PROFILE and item.get('score') not in (None,''):
                    try: target['Match Score']=str(max(0,min(100,int(item.get('score',0)))))
                    except Exception: pass
                for src_key,dst_key in mapping.items():
                    value=_clean_text(item.get(src_key,''))
                    if value: target[dst_key]=value
                _apply_local_metadata(target)
    except Exception as e:
        GROQ_CONNECTED=False; print(f'⚠️ خطا در تحلیل/استانداردسازی با Groq: {e}')
    finally:
        if http_client is not None:
            try: http_client.close()
            except Exception: pass
    return jobs


def _db_connect():
    conn=sqlite3.connect(MARKET_DB_FILE); conn.execute('PRAGMA journal_mode=WAL'); return conn


def init_market_database():
    try:
        with _db_connect() as conn:
            conn.execute("""CREATE TABLE IF NOT EXISTS jobs (
                link TEXT PRIMARY KEY, title TEXT, company TEXT, source TEXT, first_seen TEXT, last_seen TEXT,
                latest_date TEXT, city TEXT, work_mode TEXT, employment_type TEXT, salary TEXT, job_family TEXT,
                seniority TEXT, skills TEXT, match_score INTEGER, description TEXT
            )""")
            conn.execute("""CREATE TABLE IF NOT EXISTS job_observations (
                id INTEGER PRIMARY KEY AUTOINCREMENT, observed_at TEXT NOT NULL, link TEXT NOT NULL,
                title TEXT, company TEXT, source TEXT, date_text TEXT, city TEXT, work_mode TEXT,
                employment_type TEXT, salary TEXT, job_family TEXT, seniority TEXT, skills TEXT,
                match_score INTEGER, description_hash TEXT
            )""")
            conn.execute('CREATE INDEX IF NOT EXISTS idx_job_observations_observed_at ON job_observations(observed_at)')
            conn.execute("""CREATE TABLE IF NOT EXISTS scan_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT, scanned_at TEXT NOT NULL, site TEXT NOT NULL, job_count INTEGER DEFAULT 0
            )""")
        return True
    except Exception as e:
        print(f'⚠️ راه‌اندازی پایگاه داده بازار ناموفق بود: {e}'); return False


def record_market_observations(jobs):
    if not jobs: return 0
    init_market_database(); observed_at=datetime.now(timezone.utc).isoformat(); count=0
    try:
        with _db_connect() as conn:
            site_counts=Counter()
            for job in jobs:
                link=_clean_text(job.get('Link')); title=_clean_text(job.get('Title'))
                if not link or not title: continue
                _apply_local_metadata(job); score=None
                try:
                    if str(job.get('Match Score') or '').strip(): score=int(job.get('Match Score'))
                except Exception: pass
                description=_clean_text(job.get('Description','')); desc_hash=hashlib.sha256(description.encode('utf-8')).hexdigest() if description else ''
                conn.execute("""INSERT INTO jobs(link,title,company,source,first_seen,last_seen,latest_date,city,work_mode,employment_type,salary,job_family,seniority,skills,match_score,description)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(link) DO UPDATE SET
                    title=COALESCE(NULLIF(excluded.title,''),jobs.title), company=COALESCE(NULLIF(excluded.company,''),jobs.company), source=COALESCE(NULLIF(excluded.source,''),jobs.source),
                    last_seen=excluded.last_seen, latest_date=COALESCE(NULLIF(excluded.latest_date,''),jobs.latest_date), city=COALESCE(NULLIF(excluded.city,''),jobs.city),
                    work_mode=COALESCE(NULLIF(excluded.work_mode,''),jobs.work_mode), employment_type=COALESCE(NULLIF(excluded.employment_type,''),jobs.employment_type), salary=COALESCE(NULLIF(excluded.salary,''),jobs.salary),
                    job_family=COALESCE(NULLIF(excluded.job_family,''),jobs.job_family), seniority=COALESCE(NULLIF(excluded.seniority,''),jobs.seniority), skills=COALESCE(NULLIF(excluded.skills,''),jobs.skills),
                    match_score=COALESCE(excluded.match_score,jobs.match_score), description=COALESCE(NULLIF(excluded.description,''),jobs.description)""",
                    (link,title,job.get('Company',''),job.get('Source',''),observed_at,observed_at,job.get('Date',''),job.get('City',''),job.get('Work Mode',''),job.get('Employment Type',''),job.get('Salary',''),job.get('Job Family',''),job.get('Seniority',''),job.get('Skills',''),score,description))
                conn.execute("""INSERT INTO job_observations(observed_at,link,title,company,source,date_text,city,work_mode,employment_type,salary,job_family,seniority,skills,match_score,description_hash)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (observed_at,link,title,job.get('Company',''),job.get('Source',''),job.get('Date',''),job.get('City',''),job.get('Work Mode',''),job.get('Employment Type',''),job.get('Salary',''),job.get('Job Family',''),job.get('Seniority',''),job.get('Skills',''),score,desc_hash))
                site_counts[job.get('Source','نامشخص')]+=1; count+=1
            for site,site_count in site_counts.items(): conn.execute('INSERT INTO scan_events(scanned_at,site,job_count) VALUES(?,?,?)',(observed_at,site,site_count))
        return count
    except Exception as e:
        print(f'⚠️ ثبت تاریخچه بازار ناموفق بود: {e}'); return 0


def _market_iso_days_ago(days):
    return datetime.fromtimestamp(time.time()-days*86400,tz=timezone.utc).isoformat()


def generate_market_report():
    """گزارش ۳۰روزه برای داشبورد و سایت آینده؛ نماینده کل بازار کار نیست."""
    init_market_database(); start_iso=_market_iso_days_ago(MARKET_REPORT_DAYS); current_iso=_market_iso_days_ago(MARKET_TREND_WINDOW_DAYS); previous_iso=_market_iso_days_ago(MARKET_TREND_WINDOW_DAYS*2)
    report={'generated_at':datetime.now(timezone.utc).isoformat(),'window_days':MARKET_REPORT_DAYS,'methodology':'فقط بر مبنای مشاهدات منابع ثبت‌شده؛ نماینده کل بازار کار نیست.','metrics':{},'sources':[],'top_skills':[],'top_job_families':[],'top_cities':[],'top_companies':[],'salary_statistics':{},'top_skill_pairs':[],'trend':{},'validation':{}}
    try:
        with _db_connect() as conn:
            total_obs=conn.execute('SELECT COUNT(*) FROM job_observations WHERE observed_at>=?',(start_iso,)).fetchone()[0]
            unique_jobs=conn.execute('SELECT COUNT(DISTINCT link) FROM job_observations WHERE observed_at>=?',(start_iso,)).fetchone()[0]
            salary_count=conn.execute("SELECT COUNT(DISTINCT link) FROM job_observations WHERE observed_at>=? AND TRIM(COALESCE(salary,''))<>''",(start_iso,)).fetchone()[0]
            report['metrics']={'observations':total_obs,'unique_jobs':unique_jobs,'salary_coverage_percent':round(salary_count*100/unique_jobs,1) if unique_jobs else 0}
            report['sources']=[{'source':r[0],'observations':r[1]} for r in conn.execute('SELECT source,COUNT(*) FROM job_observations WHERE observed_at>=? GROUP BY source ORDER BY COUNT(*) DESC',(start_iso,)).fetchall()]
            report['top_job_families']=[{'name':r[0],'count':r[1]} for r in conn.execute("SELECT job_family,COUNT(DISTINCT link) FROM job_observations WHERE observed_at>=? AND TRIM(COALESCE(job_family,''))<>'' GROUP BY job_family ORDER BY COUNT(DISTINCT link) DESC LIMIT 20",(start_iso,)).fetchall()]
            report['top_cities']=[{'name':r[0],'count':r[1]} for r in conn.execute("SELECT city,COUNT(DISTINCT link) FROM job_observations WHERE observed_at>=? AND TRIM(COALESCE(city,''))<>'' GROUP BY city ORDER BY COUNT(DISTINCT link) DESC LIMIT 20",(start_iso,)).fetchall()]
            skill_counter=Counter()
            for (skills,) in conn.execute("SELECT skills FROM job_observations WHERE observed_at>=? AND TRIM(COALESCE(skills,''))<>''",(start_iso,)).fetchall():
                for skill in str(skills).split(','):
                    skill=_clean_text(skill)
                    if skill: skill_counter[skill]+=1
            report['top_skills']=[{'name':n,'count':c} for n,c in skill_counter.most_common(30)]
            current_count=conn.execute('SELECT COUNT(DISTINCT link) FROM job_observations WHERE observed_at>=?',(current_iso,)).fetchone()[0]
            previous_count=conn.execute('SELECT COUNT(DISTINCT link) FROM job_observations WHERE observed_at>=? AND observed_at<?',(previous_iso,current_iso)).fetchone()[0]
            report['trend']={'window_days':MARKET_TREND_WINDOW_DAYS,'current_unique_jobs':current_count,'previous_unique_jobs':previous_count,'percent_change':None if previous_count==0 else round((current_count-previous_count)*100/previous_count,1),'note':'برای تفسیر روند، داده کافی در هر دو بازه لازم است.'}
            report['top_companies']=[{'name':r[0],'count':r[1]} for r in conn.execute("SELECT company,COUNT(DISTINCT link) FROM job_observations WHERE observed_at>=? AND TRIM(COALESCE(company,''))<>'' GROUP BY company ORDER BY COUNT(DISTINCT link) DESC LIMIT 30",(start_iso,)).fetchall()]
            salary_rows=conn.execute("SELECT salary FROM job_observations WHERE observed_at>=? AND TRIM(COALESCE(salary,''))<>''",(start_iso,)).fetchall()
            salary_mins=[]; salary_maxs=[]
            for (salary_text,) in salary_rows:
                low,high,unit=_parse_salary_range(salary_text)
                if low is not None:
                    salary_mins.append(low); salary_maxs.append(high)
            report['salary_statistics']={'unit':'میلیون تومان','sample_count':len(salary_mins),'min_observed':min(salary_mins) if salary_mins else None,'max_observed':max(salary_maxs) if salary_maxs else None,'median_min':round(_median(salary_mins),2) if salary_mins else None,'median_max':round(_median(salary_maxs),2) if salary_maxs else None,'average_min':round(sum(salary_mins)/len(salary_mins),2) if salary_mins else None,'average_max':round(sum(salary_maxs)/len(salary_maxs),2) if salary_maxs else None,'note':'فقط حقوقی که مقدار و واحد آن از متن آگهی قابل استخراج بود در این آمار آمده است.'}
            pair_counter=Counter()
            for (skills,) in conn.execute("SELECT skills FROM job_observations WHERE observed_at>=? AND TRIM(COALESCE(skills,''))<>''",(start_iso,)).fetchall():
                skill_list=sorted(set(_clean_text(s) for s in str(skills).split(',') if _clean_text(s)))
                for i in range(len(skill_list)):
                    for j in range(i+1,len(skill_list)):
                        pair_counter[(skill_list[i],skill_list[j])]+=1
            report['top_skill_pairs']=[{'skill_1':a,'skill_2':b,'count':c} for (a,b),c in pair_counter.most_common(30)]
            try:
                project_start=datetime.fromisoformat(_load_state().get('project_started_at'))
                project_day=min(max(1,(datetime.now(timezone.utc)-project_start).days+1),PROJECT_VALIDATION_DAYS)
            except Exception:
                project_day=1
            report['validation']={'project_day':project_day,'project_days_target':PROJECT_VALIDATION_DAYS,'purpose':'اعتبارسنجی کیفیت داده و سنجش willingness-to-pay در ماه اول'}
            recent=conn.execute("""SELECT o.link,o.title,o.company,o.source,o.date_text,o.city,o.work_mode,o.employment_type,o.salary,o.job_family,o.seniority,o.skills,o.match_score FROM job_observations o JOIN (SELECT link,MAX(id) AS max_id FROM job_observations WHERE observed_at>=? GROUP BY link) latest ON latest.max_id=o.id ORDER BY o.observed_at DESC""",(start_iso,)).fetchall()
        with open(MARKET_DATA_CSV,'w',newline='',encoding='utf-8-sig') as f:
            writer=csv.writer(f); writer.writerow(['لینک','عنوان شغلی','شرکت','منبع','تاریخ آگهی','شهر','نوع کار','نوع همکاری','حقوق','خانواده شغلی','سطح ارشدیت','مهارت‌ها','امتیاز سازگاری']); writer.writerows(recent)
        Path(MARKET_REPORT_FILE).write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        public_payload={k:report.get(k) for k in ('generated_at','window_days','methodology','metrics','sources','top_skills','top_job_families','top_cities','top_companies','salary_statistics','top_skill_pairs','trend','validation')}
        Path(MARKET_PUBLIC_JSON).write_text(json.dumps(public_payload,ensure_ascii=False,indent=2),encoding='utf-8')
        print(f'🌐 داده عمومی و API-ready ساخته شد: {MARKET_PUBLIC_JSON}')
        with open(MARKET_SUMMARY_CSV,'w',newline='',encoding='utf-8-sig') as f:
            writer=csv.writer(f); writer.writerow(['بخش','نام','تعداد/مقدار'])
            for row in report['sources']: writer.writerow(['منبع',row['source'],row['observations']])
            for row in report['top_job_families']: writer.writerow(['خانواده شغلی',row['name'],row['count']])
            for row in report['top_cities']: writer.writerow(['شهر',row['name'],row['count']])
            for row in report['top_skills']: writer.writerow(['مهارت',row['name'],row['count']])
            writer.writerow(['کل مشاهدات','',report['metrics']['observations']]); writer.writerow(['شغل یکتای ۳۰روزه','',report['metrics']['unique_jobs']]); writer.writerow(['پوشش حقوق (%)','',report['metrics']['salary_coverage_percent']])
        print(f'📈 گزارش بازار {MARKET_REPORT_DAYS} روزه ساخته شد: {MARKET_REPORT_FILE}'); print(f'🧩 دیتاست تحلیلی آماده شد: {MARKET_DATA_CSV}')
        return report
    except Exception as e:
        print(f'⚠️ ساخت گزارش بازار ناموفق بود: {e}'); return report


def _load_state():
    if not os.path.exists(STATE_FILE):
        return {'notified_links': [], 'last_scan': {}, 'project_started_at': datetime.now(timezone.utc).isoformat(), 'site_health': {}}
    try:
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)
            if not isinstance(state, dict): raise ValueError('invalid state')
            state.setdefault('notified_links', []); state.setdefault('last_scan', {}); state.setdefault('site_health', {}); state.setdefault('project_started_at', datetime.now(timezone.utc).isoformat())
            return state
    except Exception:
        return {'notified_links': [], 'last_scan': {}, 'project_started_at': datetime.now(timezone.utc).isoformat(), 'site_health': {}}


def _save_state(state):
    try:
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except OSError as e:
        print(f"⚠️ خطا در ذخیره وضعیت: {e}")


def send_rubika_message(text):
    """ارسال اعلان به روبیکا با Bot API؛ در صورت نبود تنظیمات، هیچ کاری نمی‌کند."""
    if not RUBIKA_BOT_TOKEN or not RUBIKA_CHAT_ID:
        return False
    try:
        url = f"{RUBIKA_API_BASE}/{RUBIKA_BOT_TOKEN}/sendMessage"
        response = robust_session.post(
            url, json={'chat_id': RUBIKA_CHAT_ID, 'text': text},
            headers={'Content-Type': 'application/json', 'User-Agent': USER_AGENT},
            timeout=30, proxies=GLOBAL_PROXIES if GLOBAL_PROXIES else None,
        )
        response.raise_for_status()
        body = response.json()
        if isinstance(body, dict) and (body.get('status') == 'OK' or isinstance(body.get('data'), dict)):
            return True
        print(f"⚠️ پاسخ غیرمنتظره از روبیکا: {body}")
    except Exception as e:
        print(f"⚠️ ارسال اعلان روبیکا ناموفق بود: {e}")
    return False


def notify_good_jobs(jobs, state):
    """ارسال فقط فرصت‌های جدید و بالاتر از آستانه به روبیکا."""
    if not RUBIKA_BOT_TOKEN or not RUBIKA_CHAT_ID:
        return
    notified = set(state.get('notified_links', []))
    for job in jobs:
        try: score = int(job.get('Match Score') or 0)
        except ValueError: score = 0
        if score < RUBIKA_MIN_SCORE or job.get('Link') in notified:
            continue
        text = (f"🔔 فرصت شغلی مناسب\n\nعنوان: {job.get('Title','')}\n"
                f"شرکت: {job.get('Company','') or 'نامشخص'}\n"
                f"نام مسئول/تماس: {job.get('Contact','') or 'نامشخص'}\n"
                f"امتیاز سازگاری: {job.get('Match Score','')}/100\n"
                f"منبع: {job.get('Source','')}\n"
                f"تاریخ آگهی: {job.get('Date','') or 'نامشخص'}\n"
                f"لینک: {job.get('Link','')}")
        if send_rubika_message(text):
            notified.add(job.get('Link'))
    state['notified_links'] = list(notified)[-2000:]

def send_scan_payload_to_brain(jobs):
    """ارسال داده چشم به مغز؛ Dry-run برای تست بدون تغییر سرور واقعی."""
    if not jobs:
        return False
    payload = {
        'sent_at': datetime.now(timezone.utc).isoformat(),
        'source': 'job-research-eye',
        'schema_version': 1,
        'jobs': jobs,
    }
    if BRAIN_DRY_RUN:
        try:
            Path(BRAIN_PAYLOAD_FILE).write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8'
            )
            print(f"🧪 حالت تست مغز فعال است؛ payload در {BRAIN_PAYLOAD_FILE} ذخیره شد.")
            return True
        except OSError as e:
            print(f"⚠️ ذخیره payload تست مغز ناموفق بود: {e}")
            return False
    if not SEND_SCAN_TO_BRAIN or not BRAIN_API_URL:
        return False
    headers = {'Content-Type': 'application/json'}
    if BRAIN_API_TOKEN:
        headers['Authorization'] = f'Bearer {BRAIN_API_TOKEN}'
    try:
        response = requests.post(BRAIN_API_URL, json=payload, headers=headers, timeout=BRAIN_TIMEOUT)
        response.raise_for_status()
        print(f"🧠 داده {len(jobs)} آگهی به سرور مغز ارسال شد.")
        return True
    except Exception as e:
        print(f"⚠️ ارسال داده به سرور مغز ناموفق بود: {e}")
        return False

def get_active_target_sites():
    """انتخاب منابع فعال از تنظیمات .env؛ منابعی که فعال نشده‌اند اسکن نمی‌شوند."""
    flags = {
        "IranTalent": ENABLE_IRANTALENT,
        "Quera": ENABLE_QUERA,
        "Karboom": ENABLE_KARBOOM,
    }
    active = []
    wanted = None
    if SOURCE_SCAN_MODE not in ("all", "", "default"):
        wanted = {item.strip().lower() for item in SOURCE_SCAN_MODE.split(",") if item.strip()}
    for site in TARGET_SITES:
        if not site.get("enabled", True):
            continue
        flag = site.get("env_flag")
        if flag and not flags.get(site["name"], True):
            continue
        if wanted is not None and site["name"].lower() not in wanted:
            continue
        active.append(site)
    if MAX_SOURCES_PER_RUN > 0:
        active = active[:MAX_SOURCES_PER_RUN]
    return active


# --- اصلی ---
def main():
    print("--- شروع ربات هوشمند (نسخه پایدار، چندسایته، کم‌درخواست و آماده تفکیک چشم/مغز) ---")
    configured_scraper_proxy = get_scraper_proxies()
    if configured_scraper_proxy:
        print("🔌 مسیر اسکرپر: پروکسی صریح تنظیم‌شده")
    elif SCRAPER_USE_ENV_PROXY:
        print("🔌 مسیر اسکرپر: Proxyهای محیط سیستم فعال")
    else:
        print("🔌 مسیر اسکرپر: اتصال مستقیم؛ Proxy محیط سیستم عمداً غیرفعال است")
    print(f"🧩 نقش ربات: {BOT_ROLE}")
    if GROQ_KEY:
        print(f"🧠 مدل Groq تنظیم‌شده: {GROQ_ACTIVE_MODEL}")
    print(f"⏱️ حداقل فاصله اسکن هر سایت: {MIN_SITE_INTERVAL_SECONDS}s | Circuit Breaker: {CIRCUIT_FAILURE_THRESHOLD} خطا / {CIRCUIT_COOLDOWN_SECONDS}s")
    print("🧪 حالت مغز: " + ("فقط ذخیره payload محلی" if BRAIN_DRY_RUN else ("ارسال به سرور" if SEND_SCAN_TO_BRAIN and BRAIN_API_URL else "غیرفعال")))

    state = _load_state()
    if os.getenv('CHECK_INTERNET_BEFORE_SCAN', '0').lower() in ('1', 'true', 'yes') and not test_internet():
        return

    groq_attempts = 0
    pending_ai_jobs = []

    def maybe_retry_groq(stage_name):
        nonlocal groq_attempts
        if BOT_ROLE == 'eye' or not GROQ_KEY or GROQ_CONNECTED or groq_attempts >= GROQ_TOTAL_ATTEMPTS:
            return
        groq_attempts += 1
        print(f"🧠 تلاش Groq در مرحله «{stage_name}»...")
        test_ai_connection(groq_attempts)

    if BOT_ROLE != 'eye' and GROQ_KEY:
        maybe_retry_groq('شروع ربات')
    elif BOT_ROLE == 'eye':
        print("👁️ حالت Eye فعال است؛ تحلیل AI محلی و اعلان محلی انجام نمی‌شود.")

    active_sites = get_active_target_sites()
    if not active_sites:
        print("❌ هیچ منبع فعالی برای اسکن تنظیم نشده است.")
        return
    print("🌐 منابع فعال: " + "، ".join(site['name'] for site in active_sites))

    existing_links = {row['Link'] for row in load_existing_jobs() if row.get('Link')}
    all_jobs = []
    all_scanned_jobs = []

    for index, site in enumerate(active_sites):
        print(f"🔍 در حال اسکن: {site['name']}...")
        previous_attempt = state.get('site_health', {}).get(site['name'], {}).get('last_attempt_epoch')
        jobs = scrape_job_site(site, state)
        current_attempt = state.get('site_health', {}).get(site['name'], {}).get('last_attempt_epoch')
        was_attempted = current_attempt != previous_attempt

        all_scanned_jobs.extend(jobs)
        new_jobs = [job for job in jobs if job.get('Link') not in existing_links]
        if new_jobs:
            print(f"✅ {len(new_jobs)} مورد جدید پیدا شد.")
            all_jobs.extend(new_jobs)
            existing_links.update(job.get('Link') for job in new_jobs if job.get('Link'))
            if BOT_ROLE != 'eye':
                pending_ai_jobs.extend(new_jobs)
        else:
            print("ℹ️ مورد جدیدی پیدا نشد.")

        if was_attempted:
            state['last_scan'][site['name']] = datetime.now(timezone.utc).isoformat()
        _save_state(state)

        # حداکثر سه تلاش: اول شروع ربات، دوم بعد از Jobinja، سوم بعد از E-Estekhdam.
        if BOT_ROLE != 'eye':
            maybe_retry_groq(f"بعد از {site['name']}")

            # رکوردهایی که هنوز امتیاز ندارند نیز در صف می‌آیند تا بعداً تکمیل شوند.
            if AI_BACKFILL_LIMIT > 0:
                for scanned_job in jobs:
                    if not scanned_job.get('Match Score') and scanned_job not in pending_ai_jobs:
                        if len(pending_ai_jobs) >= AI_BACKFILL_LIMIT:
                            break
                        pending_ai_jobs.append(scanned_job)

            if GROQ_CONNECTED and pending_ai_jobs:
                print(f"🤖 تحلیل {len(pending_ai_jobs)} آگهی با Groq...")
                before_scores = sum(1 for job in pending_ai_jobs if job.get('Match Score'))
                analyzed = analyze_jobs_with_ai(pending_ai_jobs)
                after_scores = sum(1 for job in analyzed if job.get('Match Score'))
                # فقط مواردی که واقعاً امتیاز گرفته‌اند از صف حذف شوند.
                pending_ai_jobs = [job for job in analyzed if not job.get('Match Score')]
                if after_scores == before_scores and not GROQ_CONNECTED:
                    print("ℹ️ صف تحلیل Groq حفظ شد تا تلاش بعدی آن را تکمیل کند.")

        if index < len(active_sites) - 1:
            delay = random.uniform(MIN_DELAY_BETWEEN_SITES, MAX_DELAY_BETWEEN_SITES)
            print(f"⏳ فاصله بین درخواست‌ها: {delay:.1f} ثانیه")
            time.sleep(delay)

    if all_scanned_jobs:
        if pending_ai_jobs and GROQ_CONNECTED and BOT_ROLE != 'eye':
            analyzed = analyze_jobs_with_ai(pending_ai_jobs)
            pending_ai_jobs = [job for job in analyzed if not job.get('Match Score')]
        if BOT_ROLE == 'eye':
            send_scan_payload_to_brain(all_scanned_jobs)
            save_and_sort_data(all_scanned_jobs)
            print("👁️ اسکن تمام شد؛ داده خام برای مغز آماده/ارسال شد.")
        else:
            if SEND_SCAN_TO_BRAIN:
                send_scan_payload_to_brain(all_scanned_jobs)
            # تمام آگهی‌های دیده‌شده ذخیره می‌شوند تا Company/Date/Contact و امتیازهای بعدی
            # بتوانند رکوردهای قبلی را هم تکمیل کنند. اعلان فقط برای all_jobs (جدیدها) است.
            save_and_sort_data(all_scanned_jobs)
            notify_good_jobs(all_jobs, state)
            _save_state(state)
            print("✨ عملیات با موفقیت تمام شد.")
    else:
        print("❌ هیچ آگهی قابل پردازشی از سایت‌ها به دست نیامد.")

if __name__ == "__main__":
    main()
