import requests
from bs4 import BeautifulSoup
import csv
import os
import openai  # کتابخانه هوش مصنوعی
from dotenv import load_dotenv # اضافه شده برای امنیت

# --- تنظیمات ---
load_dotenv() # بارگذاری کلید از فایل .env
CSV_FILE = 'results.csv'
# فراخوانی امن کلید (بجای نوشتنِ مستقیم در کد)
groq_api_key = os.getenv("GROQ_API_KEY") 

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
}

# لیست سایت‌ها
TARGET_SITES = [
    {"name": "Jobinja", "url": "https://jobinja.ir/jobs", "selector": "h2.c-jobListView__title"},
    {"name": "E-Estekhdam", "url": "https://e-estekhdam.com/search", "selector": "div.title a"},
    {"name": "JobVision", "url": "https://jobvision.ir/jobs", "selector": "h3.job-title"}
]

# --- ۱. تست اینترنت ---
def test_internet():
    try:
        requests.get("https://www.google.com", timeout=5)
        print("✅ تست اینترنت: موفق")
        return True
    except:
        print("❌ تست اینترنت: ناموفق! VPN را چک کنید.")
        return False

# --- ۲. تست اتصال هوش مصنوعی ---
def test_ai_connection():
    print("در حال تست اتصال به سرورهای Groq...")
    if not API_KEY:
        print("⚠️ هشدار: کلید API در فایل .env پیدا نشد! لطفا فایل .env را چک کنید.")
        return False
    try:
        # اتصال به Groq با استفاده از ساختار OpenAI
        client = openai.OpenAI(
            api_key=API_KEY,
            base_url="https://api.groq.com/openai/v1"
        )
        # یک تست سبک
        client.models.list()
        print("✅ اتصال به Groq: برقرار است.")
        return True
    except Exception as e:
        print(f"⚠️ خطای اتصال به Groq: {e}")
        return False

# --- ۳. استخراج داده ---
def scrape_job_site(site_info):
    extracted_data = []
    try:
        response = requests.get(site_info['url'], headers=HEADERS, timeout=10)
        soup = BeautifulSoup(response.content, 'html.parser')
        items = soup.select(site_info['selector'])
        for item in items:
            title = item.text.strip()
            # این خط برای استخراج لینک کمی ظریف‌تر است، در کد شما همین بود
            link = item.find('a')['href'] if item.name != 'a' else item['href']
            # اصلاح لینک‌های نسبی
            if not link.startswith('http'): 
                link = "https://" + site_info['url'].split('/')[2] + link
            extracted_data.append([title, link, site_info['name']])
    except Exception as e:
        print(f"خطا در سایت {site_info['name']}: {e}")
    return extracted_data

# --- ۴. ذخیره و مرتب‌سازی ---
def save_and_sort_data(new_data):
    all_data = []
    if os.path.exists(CSV_FILE):
        with open(CSV_FILE, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            next(reader, None)
            all_data = list(reader)
    
    all_data.extend(new_data)
    # حذف تکراری‌ها
    unique_data = {row[1]: row for row in all_data}.values()
    # مرتب‌سازی بر اساس نام
    sorted_data = sorted(list(unique_data), key=lambda x: x[0])
    
    with open(CSV_FILE, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Title', 'Link', 'Source'])
        writer.writerows(sorted_data)
    print(f"💾 فایل با {len(sorted_data)} مورد به‌روز شد.")

# --- اجرای نهایی ---
def main():
    print("--- شروع رباتِ هوشمند ---")
    
    if not test_internet(): return
    if not test_ai_connection(): return

    all_jobs = []
    for site in TARGET_SITES:
        print(f"در حال پردازش: {site['name']}...")
        all_jobs.extend(scrape_job_site(site))
    
    if all_jobs:
        save_and_sort_data(all_jobs)
        print("عملیات با موفقیت پایان یافت.")
    else:
        print("داده‌ای یافت نشد.")

if __name__ == "__main__":
    main()
