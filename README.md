# Job Alert Automation (Python + Telegram)

Scans configured company career pages, detects likely Data Analyst / Business Analyst / Power BI / BI roles, sends Telegram alerts for new matches, and records results in `job_log.csv`. Deduplication state is stored in `seen_jobs.json`.

## Important coverage notes

- **Official career pages:** Best coverage comes from a correct career URL and, where available, an ATS type and slug. Greenhouse and Lever public job-board endpoints are supported. Generic HTML pages are scanned for links whose visible text resembles a target role; JavaScript-heavy sites may need a site-specific adapter.
- **LinkedIn:** Do not scrape LinkedIn or automate a logged-in browser. Configure native LinkedIn job alerts and have their emails delivered to a mailbox that this script can read via IMAP.
- **Naukri:** Configure native Naukri job alerts and deliver those emails to the configured mailbox. Email parsing is heuristic; templates can change, so validate it with real alert messages.
- This is best-effort monitoring, not a guarantee that every vacancy will be detected. It does not bypass logins, CAPTCHAs, robots policies, or site access controls.

## 1. Prepare your companies

Edit `companies.csv` in Excel or a text editor. Keep the columns:

`company,careers_url,ats,ats_slug,active`

- `company`: official company name
- `careers_url`: public careers page URL
- `ats`: `greenhouse`, `lever`, or `generic`
- `ats_slug`: Greenhouse/Lever board slug (often the last URL segment)
- `active`: `1` to scan, `0` to skip

Replace the example rows with your 200 companies. A generic URL can work, but a company may use a separate job platform or dynamic page. Verify the URL before enabling it.

## 2. Create your Telegram bot

1. In Telegram, open **@BotFather** and run `/newbot`.
2. Follow the prompts and copy the bot token.
3. Open your new bot and press **Start** (or send `/start`).
4. Get your chat ID by opening `https://api.telegram.org/bot<YOUR_TOKEN>/getUpdates` in a browser after messaging the bot. Find `message.chat.id`.
5. Keep the token private. Never commit it to GitHub or include it in `companies.csv`.

## 3. Configure LinkedIn and Naukri email alerts (optional but recommended)

Create native job alerts for:
- Data Analyst
- Business Analyst
- Power BI Developer / Power BI Analyst / BI Developer

Use India-wide or your preferred locations and 1–3 years of experience where the platform permits. Route those emails to the mailbox used for this automation. For Gmail, enable IMAP and create an app password if your account supports it. Put the IMAP settings in environment variables / GitHub Secrets. Do not hardcode credentials.

Email alert parsing is best-effort. Email formats differ; initially test it against actual alert emails and check the workflow logs.

## 4. Run locally for a test

Requires Python 3.10+.

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
```

Set environment variables in your terminal (or use a local `.env` loader if you add one). At minimum set `TELEGRAM_BOT_TOKEN` and `TELEGRAM_CHAT_ID`, then run:

```bash
python main.py
```

The script writes `job_log.csv` and `seen_jobs.json` in the project directory.

## 5. Deploy to GitHub Actions (cloud, every 30 minutes)

1. Create a **private** GitHub repository.
2. Upload this project, including `.github/workflows/job-alerts.yml`.
3. In **Settings → Secrets and variables → Actions → New repository secret**, add:
   - `TELEGRAM_BOT_TOKEN`
   - `TELEGRAM_CHAT_ID`
   - Optional email secrets: `IMAP_HOST`, `IMAP_USER`, `IMAP_APP_PASSWORD`, `IMAP_FOLDER`
4. Ensure the workflow file is on the `main` branch and Actions are enabled.
5. Run **Actions → Job Alert Scanner → Run workflow** once to test.
6. Check the run logs and the committed `job_log.csv` / `seen_jobs.json`.

GitHub scheduled workflows may be delayed during periods of high load, and they may be disabled in inactive repositories. The `*/30 * * * *` schedule is UTC-based and nominally runs every 30 minutes; it is not a strict real-time SLA.

### Privacy note
The workflow commits detected job metadata and deduplication hashes to the private repository. Do not add mailbox contents, bot tokens, passwords, or other secrets to the CSV or repository. If you don't want the job log committed to GitHub, change the persistence design before deployment.

## Matching behavior

- Role titles are matched against common Data Analyst, Business Analyst, Power BI, BI, reporting, MIS, and analytics variants.
- Listings explicitly requiring more than 3 years are excluded when detected. Listings with no readable experience requirement are retained as “not specified” to reduce missed opportunities.
- This does not infer seniority perfectly. Review results and tune `ROLE_TERMS`, `EXCLUDE_TERMS`, and the experience parser in `main.py`.
- Duplicate identity is based on company + URL (query parameters removed).

## Known limitations / next improvements

- Generic career-page scanning only sees links in returned HTML; it may miss JavaScript-rendered postings.
- Some pages block automated requests or have non-standard structures. Add an ATS-specific adapter for those sites.
- LinkedIn/Naukri email formats may not consistently expose the actual role title as anchor text. Validate real messages and adjust `scan_alert_emails()` if needed.
- For production reliability, consider a small database (PostgreSQL/SQLite on persistent storage) instead of committing state to Git.
