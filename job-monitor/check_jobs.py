#!/usr/bin/env python3
"""
Job monitor: checks target companies for new product manager postings
and sends an email when new roles are found.

Usage:
  python check_jobs.py            # normal run — check + email if new jobs
  python check_jobs.py --seed     # baseline run — record all current jobs, no email
"""

import json
import os
import smtplib
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import requests

SCRIPT_DIR = Path(__file__).parent
COMPANIES_FILE = SCRIPT_DIR / "companies.json"
SEEN_JOBS_FILE = SCRIPT_DIR / "seen_jobs.json"

NOTIFY_EMAIL = os.environ.get("NOTIFY_EMAIL", "ngardone@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")


# ---------------------------------------------------------------------------
# State helpers
# ---------------------------------------------------------------------------

def load_seen_ids() -> set:
    if SEEN_JOBS_FILE.exists():
        with open(SEEN_JOBS_FILE) as f:
            return set(json.load(f))
    return set()


def save_seen_ids(ids: set) -> None:
    with open(SEEN_JOBS_FILE, "w") as f:
        json.dump(sorted(ids), f, indent=2)
        f.write("\n")


# ---------------------------------------------------------------------------
# Fetchers
# ---------------------------------------------------------------------------

def fetch_greenhouse(company: dict) -> list[dict]:
    token = company["board_token"]
    url = f"https://boards-api.greenhouse.io/v1/boards/{token}/jobs"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    jobs = []
    keyword = company["filter"].lower()
    for job in resp.json().get("jobs", []):
        title = job.get("title", "")
        if keyword in title.lower():
            jobs.append({
                "id": f"greenhouse-{job['id']}",
                "company": company["name"],
                "title": title,
                "url": job.get("absolute_url", ""),
            })
    return jobs


def fetch_workday(company: dict) -> list[dict]:
    # Workday caps at 20 results per request; paginate to collect all matching jobs.
    PAGE_SIZE = 20
    base_url = company.get("base_url", "")
    keyword = company["filter"].lower()
    jobs = []
    offset = 0

    while True:
        payload = {
            "appliedFacets": {},
            "limit": PAGE_SIZE,
            "offset": offset,
            "searchText": company.get("search_text", "product"),
        }
        resp = requests.post(
            company["api_url"],
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        postings = data.get("jobPostings", [])

        for job in postings:
            title = job.get("title", "")
            if keyword in title.lower():
                external_path = job.get("externalPath", "")
                job_id = external_path.split("_")[-1] if "_" in external_path else external_path
                jobs.append({
                    "id": f"workday-{job_id}",
                    "company": company["name"],
                    "title": title,
                    "url": base_url + external_path if external_path else "",
                })

        total = data.get("total", 0)
        offset += PAGE_SIZE
        if offset >= total or not postings:
            break

    return jobs


def fetch_ashby(company: dict) -> list[dict]:
    org_name = company["org_name"]  # e.g. "Superhuman Platform Inc"
    keyword = company["filter"].lower()
    org_slug = org_name.replace(" ", "%20")

    payload = {
        "operationName": "ApiJobBoardWithTeams",
        "variables": {"organizationHostedJobsPageName": org_name},
        "query": """
        query ApiJobBoardWithTeams($organizationHostedJobsPageName: String!) {
          jobBoard: jobBoardWithTeams(
            organizationHostedJobsPageName: $organizationHostedJobsPageName
          ) {
            jobPostings {
              id title teamId locationId locationName
            }
          }
        }
        """,
    }
    resp = requests.post(
        "https://jobs.ashbyhq.com/api/non-user-graphql",
        json=payload,
        headers={
            "Content-Type": "application/json",
            "Referer": f"https://jobs.ashbyhq.com/{org_slug}?embed=js",
        },
        timeout=30,
    )
    resp.raise_for_status()

    jobs = []
    for job in resp.json()["data"]["jobBoard"]["jobPostings"]:
        title = job.get("title", "")
        if keyword in title.lower():
            jobs.append({
                "id": f"ashby-{job['id']}",
                "company": company["name"],
                "title": title,
                "url": f"https://jobs.ashbyhq.com/{org_slug}/{job['id']}",
            })
    return jobs


FETCHERS = {
    "greenhouse": fetch_greenhouse,
    "workday": fetch_workday,
    "ashby": fetch_ashby,
}


def fetch_jobs(company: dict) -> list[dict]:
    fetcher = FETCHERS.get(company["type"])
    if not fetcher:
        print(f"  [SKIP] Unknown type '{company['type']}' for {company['name']}")
        return []
    try:
        jobs = fetcher(company)
        print(f"  {len(jobs)} matching role(s) found")
        return jobs
    except Exception as exc:
        print(f"  [ERROR] {exc}")
        return []


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def send_email(new_jobs: list[dict]) -> None:
    if not GMAIL_APP_PASSWORD:
        print("[ERROR] GMAIL_APP_PASSWORD not set — skipping email")
        return

    date_str = datetime.now().strftime("%B %-d, %Y")
    subject = f"New PM Job Postings — {date_str}"

    lines = [f"New product management roles posted as of {date_str}:\n"]
    for job in new_jobs:
        lines += [
            f"Company:  {job['company']}",
            f"Role:     {job['title']}",
            f"Link:     {job['url']}",
            "",
        ]
    body = "\n".join(lines)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = NOTIFY_EMAIL
    msg["To"] = NOTIFY_EMAIL
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(NOTIFY_EMAIL, GMAIL_APP_PASSWORD)
        server.sendmail(NOTIFY_EMAIL, NOTIFY_EMAIL, msg.as_string())

    print(f"Email sent: {len(new_jobs)} new role(s) to {NOTIFY_EMAIL}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    seed_mode = "--seed" in sys.argv

    with open(COMPANIES_FILE) as f:
        companies = json.load(f)

    seen_ids = load_seen_ids()
    all_current_ids: set = set()
    new_jobs: list[dict] = []

    for company in companies:
        print(f"Checking {company['name']}...")
        jobs = fetch_jobs(company)
        for job in jobs:
            all_current_ids.add(job["id"])
            if job["id"] not in seen_ids:
                new_jobs.append(job)
                print(f"    NEW: {job['title']}")

    # Persist: union of previously seen + all currently live IDs.
    # We never remove IDs so a role that disappears and reappears won't re-notify.
    save_seen_ids(seen_ids | all_current_ids)

    if seed_mode:
        print(f"\nSeed complete — {len(all_current_ids)} roles baselined. No email sent.")
        return

    if new_jobs:
        send_email(new_jobs)
    else:
        print("No new roles found.")


if __name__ == "__main__":
    main()
