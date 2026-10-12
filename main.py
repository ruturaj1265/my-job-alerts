import csv
import hashlib
import json
import os
import re
import time
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timezone
from html import unescape
from urllib.parse import urljoin, urlparse
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import os
import re
from dotenv import load_dotenv
load_dotenv()
from concurrent.futures import ThreadPoolExecutor, as_completed

ROOT = os.path.dirname(os.path.abspath(__file__))
COMPANIES_FILE = os.path.join(ROOT, "companies.csv")
SEEN_FILE = os.path.join(ROOT, "seen_jobs.json")
LOG_FILE = os.path.join(ROOT, "job_log.csv")
TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "8"))
USER_AGENT = os.getenv("USER_AGENT", "PersonalJobAlertBot/1.0 (career-page monitoring)")
ROLE_TERMS = [
    "data analyst", "business analyst", "power bi developer",
    "power bi analyst", "bi developer", "business intelligence analyst",
    "reporting analyst", "analytics analyst", "mis analyst",
    "operations analyst", "data & business analyst", "data and business analyst"
]
EXCLUDE_TERMS = [
    "intern", "internship", "director", "vice president", "vp ",
    "head of", "chief ", "principal", "staff engineer"
]
EXPERIENCE_RE = re.compile(r"(?<!\d)([0-9]{1,2})\s*(?:-|to)\s*([0-9]{1,2})\s*(?:years?|yrs?)", re.I)
SINGLE_EXP_RE = re.compile(r"(?<!\d)([0-9]{1,2})\s*\+?\s*(?:years?|yrs?)", re.I)

def now_iso():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")

def load_seen():
    try:
        with open(SEEN_FILE, encoding="utf-8") as f:
            return set(json.load(f))
    except (OSError, json.JSONDecodeError):
        return set()

def save_seen(seen):
    with open(SEEN_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(seen), f, indent=2)

def ensure_log():
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, "w", newline="", encoding="utf-8-sig") as f:
            csv.DictWriter(f, fieldnames=[
                "detected_at", "company", "title", "location", "source",
                "posted_date", "experience", "url", "match_reason"
            ]).writeheader()

def append_log(job):
    ensure_log()
    with open(LOG_FILE, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "detected_at", "company", "title", "location", "source",
            "posted_date", "experience", "url", "match_reason"
        ])
        writer.writerow(job)

def send_telegram(text):
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        print("Telegram not configured; notification skipped.")
        return False
    endpoint = f"https://api.telegram.org/bot{token}/sendMessage"
    # Telegram message limit is 4096 characters; keep a margin.
    response = requests.post(endpoint, json={
        "chat_id": chat_id,
        "text": text[:3900],
        "disable_web_page_preview": True
    }, timeout=TIMEOUT)
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(f"Telegram API error: {payload}")
    return True



def read_companies():
    excel_file = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "Top Companies List.xlsx"
    )

    workbook = load_workbook(
        excel_file,
        read_only=False,
        data_only=False
    )

    sheet_name = "All Companies (1-600)"

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {workbook.sheetnames}"
        )

    sheet = workbook[sheet_name]

    # Read the column headers from row 3.
    headers = {
        str(cell.value).strip().lower(): cell.column
        for cell in sheet[3]
        if cell.value is not None
    }

    required = ["company name", "career portal link"]

    for header in required:
        if header not in headers:
            raise ValueError(
                f"Required column '{header}' not found. "
                f"Headers found: {list(headers.keys())}"
            )

    companies = []

    # Your screenshot shows headers on row 3 and data below.
    for row_num in range(4, sheet.max_row + 1):
        company = sheet.cell(
            row_num, headers["company name"]
        ).value

        if not company:
            continue

       
        link_cell = sheet.cell(
            row_num, headers["career portal link"]
        )

        career_url = ""

        # Case 1: Normal Excel hyperlink
        if link_cell.hyperlink:
            career_url = link_cell.hyperlink.target or ""

        # Case 2: Excel HYPERLINK() formula
        elif isinstance(link_cell.value, str):
            formula = link_cell.value.strip()

            
            match = re.search(r'https?://[^"]+', formula, re.IGNORECASE)

            if match:
                career_url = match.group(0)


            # Case 3: The cell contains a plain URL
            elif formula.startswith(("https://", "http://")):
                career_url = formula

        career_url = career_url.strip()

        if not career_url.startswith(("https://", "http://")):
            career_url = ""

        # Activate all 600 companies.
        rank_col = headers.get("rank")
        rank_value = (
            sheet.cell(row_num, rank_col).value
            if rank_col else row_num - 3
        )

        try:
            rank = int(rank_value)
        except (TypeError, ValueError):
            rank = row_num - 3

        companies.append({
            "company": str(company).strip(),
            "careers_url": career_url,
            "ats": "generic",
            "ats_slug": "",
            "active": "1"
        })

    workbook.close()

    print(f"Loaded {len(companies)} companies from Excel.")
    print(
        "Active companies:",
        sum(c["active"] == "1" for c in companies)
    )

    return companies

def norm(s):
    return re.sub(r"\s+", " ", unescape(s or "")).strip()

def role_match(title):
    t = norm(title).lower()
    if not t or any(x in t for x in EXCLUDE_TERMS):
        return None
    for term in ROLE_TERMS:
        if term in t:
            return term
    # Accept generic BI Analyst but not unrelated analyst roles.
    if re.search(r"\bbi analyst\b", t):
        return "BI analyst"
    return None

def experience_match(text):
    text = norm(text)
    ranges = EXPERIENCE_RE.findall(text)
    if ranges:
        for a, b in ranges:
            lo, hi = int(a), int(b)
            if lo <= 3 and hi >= 1:
                return True, f"{lo}-{hi} years"
        return False, ""
    # Single numbers such as "2 years experience"; avoid interpreting dates/IDs.
    matches = SINGLE_EXP_RE.findall(text)
    if matches:
        nums = [int(x) for x in matches if int(x) <= 30]
        if nums and any(1 <= n <= 3 for n in nums):
            return True, f"{nums[0]} years mentioned"
        if nums and all(n > 3 for n in nums):
            return False, ""
    # If no experience is specified, keep it as a potential match rather than
    # silently discarding a suitable job; mark the reason as unspecified.
    return None, "not specified"

def make_job(company, title, url, location="", source="", posted_date="", description=""):
    title = norm(title)
    match = role_match(title)
    if not match or not url or not url.startswith(("http://", "https://")):
        return None
    exp_ok, exp_label = experience_match(f"{title} {description}")
    if exp_ok is False:
        return None
    reason = f"Role keyword: {match}; experience: {exp_label}"
    # Stable key prevents the same listing being re-alerted if its text changes.
    key = hashlib.sha256((company.lower() + "|" + url.split("?")[0].rstrip("/")).encode()).hexdigest()
    return {
        "_key": key,
        "detected_at": now_iso(),
        "company": norm(company),
        "title": title,
        "location": norm(location) or "Not specified",
        "source": source,
        "posted_date": norm(posted_date) or "Not specified",
        "experience": exp_label,
        "url": url,
        "match_reason": reason,
    }

def fetch_json(url):
    r = requests.get(url, timeout=TIMEOUT, headers={"User-Agent": USER_AGENT})
    r.raise_for_status()
    return r.json()

def scan_greenhouse(company, slug):
    jobs = []
    data = fetch_json(f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true")
    for item in data.get("jobs", []):
        title = item.get("title", "")
        loc = (item.get("location") or {}).get("name", "")
        desc = BeautifulSoup(item.get("content", ""), "html.parser").get_text(" ", strip=True)
        job = make_job(company, title, item.get("absolute_url", ""), loc, "Company careers (Greenhouse)",
                       item.get("updated_at", ""), desc)
        if job: jobs.append(job)
    return jobs

def scan_lever(company, slug):
    jobs = []
    data = fetch_json(f"https://api.lever.co/v0/postings/{slug}?mode=json")
    for item in data if isinstance(data, list) else []:
        title = item.get("text", "")
        categories = item.get("categories") or {}
        loc = categories.get("location", "")
        desc = " ".join([item.get("descriptionPlain", ""), item.get("additionalPlain", "")])
        job = make_job(company, title, item.get("hostedUrl", ""), loc, "Company careers (Lever)",
                       item.get("createdAt", ""), desc)
        if job: jobs.append(job)
    return jobs

def scan_generic_career_page(company, url):
    jobs = []
    if not url:
        return jobs
    try:
        r = requests.get(url, timeout=TIMEOUT, headers={"User-Agent": USER_AGENT})
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        base = r.url
        for a in soup.find_all("a", href=True):
            title = norm(a.get_text(" ", strip=True) or a.get("aria-label", ""))
            href = urljoin(base, a["href"].strip())
            if not role_match(title):
                continue
            if urlparse(href).scheme not in ("http", "https"):
                continue
            # Avoid links to generic category/search pages when possible.
            parent_text = norm(a.parent.get_text(" ", strip=True))[:1200] if a.parent else ""
            job = make_job(company, title, href, "", "Company careers page", "", parent_text)
            if job: jobs.append(job)
    except Exception as exc:
        print(f"[WARN] Career page failed for {company} ({url}): {exc}")
    return jobs

def scan_company(row):
    company = norm(row.get("company"))
    ats = norm(row.get("ats", "")).lower()
    slug = norm(row.get("ats_slug", ""))
    careers_url = norm(row.get("careers_url", ""))
    try:
        if ats == "greenhouse" and slug:
            return scan_greenhouse(company, slug)
        if ats == "lever" and slug:
            return scan_lever(company, slug)
        return scan_generic_career_page(company, careers_url)
    except Exception as exc:
        print(f"[WARN] ATS scan failed for {company}: {exc}")
        return scan_generic_career_page(company, careers_url) if careers_url else []

def decode_mime(value):
    if not value: return ""
    parts = []
    for chunk, enc in decode_header(value):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(enc or "utf-8", errors="replace"))
        else:
            parts.append(chunk)
    return "".join(parts)

def email_body(msg):
    chunks = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get_content_disposition() == "attachment":
                continue
            ctype = part.get_content_type()
            if ctype in ("text/plain", "text/html"):
                raw = part.get_payload(decode=True)
                if raw:
                    chunks.append(raw.decode(part.get_content_charset() or "utf-8", errors="replace"))
    else:
        raw = msg.get_payload(decode=True)
        if raw:
            chunks.append(raw.decode(msg.get_content_charset() or "utf-8", errors="replace"))
    return "\n".join(chunks)

def scan_alert_emails():
    """Read configured job-alert mailbox via IMAP; uses read-only search and recent messages."""
    host = os.getenv("IMAP_HOST", "").strip()
    user = os.getenv("IMAP_USER", "").strip()
    password = os.getenv("IMAP_APP_PASSWORD", "").strip()
    folder = os.getenv("IMAP_FOLDER", "INBOX").strip()
    if not all([host, user, password]):
        return []
    jobs = []
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(host)
        mail.login(user, password)
        mail.select(folder, readonly=True)
        status, data = mail.search(None, "SINCE", datetime.now().strftime("%d-%b-%Y"))
        if status != "OK":
            return jobs
        ids = data[0].split()[-100:]  # cap each scan to recent 100 messages
        companies = read_companies()
        company_names = [norm(c["company"]) for c in companies]
        for msg_id in ids:
            status, parts = mail.fetch(msg_id, "(RFC822)")
            if status != "OK" or not parts or not isinstance(parts[0], tuple):
                continue
            msg = email.message_from_bytes(parts[0][1])
            subject = decode_mime(msg.get("Subject", ""))
            sender = msg.get("From", "").lower()
            if not any(x in (subject + " " + sender).lower() for x in ["linkedin", "naukri", "job alert", "jobs for you", "job mail"]):
                continue
            body = email_body(msg)
            soup = BeautifulSoup(body, "html.parser")
            text_body = soup.get_text(" ", strip=True) if "<html" in body.lower() else body
            links = []
            for a in soup.find_all("a", href=True):
                href = a["href"]
                title = norm(a.get_text(" ", strip=True))
                if href.startswith(("http://", "https://")) and title:
                    links.append((title, href))
            # Match links by visible title; source email is only a discovery path.
            for title, href in links:
                if not role_match(title):
                    continue
                lower_blob = (subject + " " + text_body).lower()
                company = next((c for c in company_names if c.lower() in lower_blob), "Company from job alert")
                source = "LinkedIn email alert" if "linkedin" in sender or "linkedin" in subject.lower() else "Naukri email alert"
                job = make_job(company, title, href, "", source, "", text_body[:1500])
                if job: jobs.append(job)
    except Exception as exc:
        print(f"[WARN] Email alert scan failed: {exc}")
    finally:
        if mail:
            try: mail.logout()
            except Exception: pass
    return jobs

def main():
    ensure_log()
    seen = load_seen()
    candidates = []
    companies = read_companies()
    print(f"Scanning {len(companies)} active companies at {now_iso()}")
    
    max_workers = int(os.getenv("SCAN_WORKERS", "8"))
    completed = 0

    print(f"Scanning with {max_workers} concurrent workers.")

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(scan_company, row): row
            for row in companies
        }

        for future in as_completed(futures):
            row = futures[future]

            try:
                candidates.extend(future.result())
            except Exception as exc:
                print(f"[WARN] {row.get('company')}: {exc}")

            completed += 1
            if completed % 50 == 0 or completed == len(companies):
                print(f"Scan progress: {completed}/{len(companies)} companies")

    candidates.extend(scan_alert_emails())

    new_jobs = []
    for job in candidates:
        key = job.pop("_key")
        if key in seen:
            continue
        seen.add(key)
        append_log(job)
        new_jobs.append(job)

    save_seen(seen)
    print(f"Found {len(candidates)} candidate matches; {len(new_jobs)} new jobs.")
    for job in new_jobs:
        
        msg = (
            "🚀 MY JOB ALERT\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "🎯 NEW JOB MATCH\n\n"
            f"🏢 Company: {job['company']}\n"
            f"💼 Role: {job['title']}\n"
            f"📍 Location: {job['location']}\n"
            f"🧑‍💻 Experience: {job['experience']}\n"
            f"🔎 Source: {job['source']}\n"
            f"📅 Posted: {job['posted_date']}\n\n"
            "✨ WHY IT MATCHES\n"
            f"✅ {job['match_reason']}\n\n"
            "🔗 APPLY NOW\n"
            f"{job['url']}\n\n"
            "⚡ Your next career move starts here"
        )

        try:
            send_telegram(msg)
        except Exception as exc:
            print(f"[WARN] Telegram notification failed for {job['title']}: {exc}")
    if not new_jobs:
        print("No new matching jobs this run.")

if __name__ == "__main__":
    main()
