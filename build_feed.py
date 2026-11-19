#!/usr/bin/env python3
"""
Master Career Database -> Daily Curated Jobs RSS Feed (v3)
Built directly from: Expanded_Master_Career_Database_Aron_Miszlivetz_v2.xlsx
Target Profile: EU Professional — Consultancies, Public Policy, EU Government Affairs Manager,
                Energy, Tech/Digital/AI Governance, Defence & Generalist EU Policy Roles.
Target Hubs: Brussels, Paris, The Hague (Den Haag), Berlin, Benelux (Luxembourg, Amsterdam,
             Rotterdam, Antwerp), Vienna, Geneva & EU-Remote.

Live Scrapers Included (Zero external dependencies — Python 3.10+ stdlib only):
  - SmartRecruiters JSON API: OECD / IEA, Statkraft, Bosch Group
  - Greenhouse JSON API:      FGS Global, Brunswick Group, Teneo, Cloudflare, Anthropic, Helsing
  - Ashby JSON API:           OpenAI, Mistral AI
  - Lever JSON API:           Future of Life Institute
  - Workable JSON API:        Ember (Energy Think Tank)
  - Jobylon HTML Scraper:     Rud Pedersen Public Affairs
  - Teamtailor RSS:           Forward Global
  - Native RSS & APIs:        EU Careers (EPSO), UN Careers, ReliefWeb API
  - Employer DB Matcher:      Auto-tags & boosts matches across all 200 CV-Tailored + 794 Priority Targets
"""

import html
import json
import os
import re
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from email.utils import format_datetime, parsedate_to_datetime

# ==============================================================================
# 1. PROFILE & DATABASE CONFIGURATION
# ==============================================================================

TARGET_LOCATIONS = [
    "brussels", "bruxelles", "brussel", "belgium", "belgique",
    "paris", "ile-de-france", "île-de-france", "france",
    "the hague", "den haag", "hague", "amsterdam", "amstelveen", "rotterdam", "utrecht", "veldhoven", "netherlands", "holland",
    "luxembourg", "benelux",
    "berlin", "bonn", "frankfurt", "munich", "düsseldorf", "dusseldorf", "germany", "deutschland",
    "vienna", "wien", "geneva", "genève", "zurich",
    "remote europe", "eu remote", "european union", "remote (anywhere)", "europe", "emea",
]

# Core policy, public affairs, energy, tech/digital & EU generalist keywords
CORE_KEYWORDS = [
    "public affairs", "government affairs", "government relations", "regulatory strategy",
    "regulatory affairs", "institutional relations", "institutional affairs", "corporate affairs",
    "eu policy", "european policy", "eu affairs", "european affairs", "public policy",
    "policy manager", "policy advisor", "policy adviser", "policy officer", "policy analyst",
    "energy security", "energy policy", "energy strategy", "energy analyst", "energy transition",
    "grid", "tso", "lng", "hydrogen", "electricity", "nuclear", "repowereu", "dg ener",
    "digital policy", "tech policy", "ai policy", "ai governance", "dsa", "dma", "nis2",
    "cybersecurity", "defence", "defense", "nato", "geopolitical", "political risk",
    "multilateral", "trilogue", "foreign affairs", "external relations", "stakeholder",
    "advocacy", "horizon europe", "eu funding", "eu tender", "human rights", "humanitarian",
    "strategic communication", "strategic communications", "generalist", "international relations",
]

# Seniority & role-level signals
SENIORITY_SIGNALS = [
    "senior", "director", "head of", "manager", "lead", "principal", "adviser", "advisor",
    "expert", "consultant", "managing consultant", "associate director", "account director",
    "vp", "vice president", "p-3", "p-4", "p-5", "ad5", "ad7", "pal5", "pal6", "a3", "a4", "fg iv",
]

# Hard exclusions (drops internships, junior roles, and unrelated technical/admin jobs)
EXCLUDE_KEYWORDS = [
    "intern ", "internship", "trainee", "traineeship", "stagiaire", "stage ", "stage -",
    "blue book", "schuman", "apprentice", "working student", "werkstudent", "formula student",
    "junior ", "entry level", "entry-level", "graduate scheme", "volunteer",
    "software engineer", "frontend", "front-end", "backend", "back-end", "fullstack", "devops",
    "data engineer", "machine learning engineer", "research scientist", "simulation engineer",
    "pentester", "it support", "technicien support", "administrateur système", "system safety",
    "accountant", "contrôleur financier", "contrôleur de gestion", "payroll", "nurse",
    "warehouse", "driver", "intraday shift trader", "shift trader", "gépmester", "gyakornok",
]

MIN_SCORE = 5  # Minimum relevance score to appear in feed.xml

# ==============================================================================
# 2. VERIFIED LIVE EMPLOYER BOARDS FROM YOUR MASTER DATABASE
# ==============================================================================

SMARTRECRUITERS_BOARDS = [
    {"name": "OECD / IEA", "slug": "OECD", "pillar": "Multilateral & Energy", "require_location": False},
    {"name": "Statkraft", "slug": "Statkraft1", "pillar": "Corporate Gov Affairs (Energy)", "require_location": True},
    {"name": "Bosch Group", "slug": "BoschGroup", "pillar": "Corporate Gov Affairs (Tech/Industrial)", "require_location": True},
]

GREENHOUSE_BOARDS = [
    {"name": "FGS Global", "slug": "fgsglobal", "pillar": "Public Affairs Consultancy", "require_location": True},
    {"name": "Brunswick Group", "slug": "brunswickgroup", "pillar": "Public Affairs Consultancy", "require_location": True},
    {"name": "Teneo", "slug": "teneo", "pillar": "Public Affairs Consultancy", "require_location": True},
    {"name": "Cloudflare", "slug": "cloudflare", "pillar": "Corporate Gov Affairs (Tech/Cyber)", "require_location": True},
    {"name": "Anthropic", "slug": "anthropic", "pillar": "Corporate Gov Affairs (AI Policy)", "require_location": True},
    {"name": "Helsing", "slug": "helsing", "pillar": "Defence AI & Security", "require_location": True},
]

ASHBY_BOARDS = [
    {"name": "OpenAI", "slug": "openai", "pillar": "Corporate Gov Affairs (AI Policy)", "require_location": True},
    {"name": "Mistral AI", "slug": "mistral.ai", "pillar": "Corporate Gov Affairs (AI Policy)", "require_location": True},
]

LEVER_BOARDS = [
    {"name": "Future of Life Institute", "slug": "futureof-life", "pillar": "Digital & AI Governance", "require_location": True},
]

WORKABLE_BOARDS = [
    {"name": "Ember", "slug": "ember-1", "pillar": "EU Energy Think Tank", "require_location": True},
]

JOBYLON_BOARDS = [
    {"name": "Rud Pedersen", "url": "https://emp.jobylon.com/companies/1663-rud-pedersen/", "pillar": "Public Affairs Consultancy", "require_location": True},
]

RSS_SOURCES = [
    {
        "name": "EU Careers (EPSO)",
        "url": "https://eu-careers.europa.eu/node/244/rss_en",
        "pillar": "EU Institutions",
        "require_location": False,
    },
    {
        "name": "UN Careers",
        "url": "https://careers.un.org/jobfeed",
        "pillar": "Multilateral / UN",
        "require_location": True,
    },
    {
        "name": "Forward Global",
        "url": "https://careers.forwardglobal.com/jobs.rss",
        "pillar": "Public Affairs Consultancy",
        "require_location": True,
    },
    # PASTE YOUR KILL-THE-NEWSLETTER FEEDS (EuroBrussels, Euractiv Jobs, Politico, LinkedIn) BELOW:
    # {
    #     "name": "EuroBrussels Daily",
    #     "url": "https://kill-the-newsletter.com/feeds/YOUR_FEED_ID.xml",
    #     "pillar": "EU Affairs Portal",
    #     "require_location": False,
    # },
]

RELIEFWEB_API_URL = (
    "https://api.reliefweb.int/v1/jobs?appname=eu-policy-rss-personal"
    "&limit=50&preset=latest"
    "&fields[include][]=title&fields[include][]=url&fields[include][]=date.created"
    "&fields[include][]=country.name&fields[include][]=city.name&fields[include][]=source.name"
    "&fields[include][]=body"
)

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; PersonalJobRSSBot/3.0)"}


# ==============================================================================
# 3. LOAD MASTER EMPLOYER DATABASE (200 CV-TAILORED + 794 PRIORITY TARGETS)
# ==============================================================================

def load_employer_db(path: str = "employer_db.json"):
    if not os.path.exists(path):
        return [], []
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("cv_200", []), data.get("priority_extra", [])


CV_200_DB, PRIORITY_EXTRA_DB = load_employer_db()


def match_database_employer(text: str):
    text_l = text.lower()
    for entry in CV_200_DB:
        key = entry["match_key"]
        if len(key) >= 4 and key in text_l:
            return entry["name"], entry["pillar"], 4  # +4 pts for CV_200 Priority Target
    for entry in PRIORITY_EXTRA_DB:
        key = entry["match_key"]
        if len(key) >= 5 and key in text_l:
            return entry["name"], entry["pillar"], 2  # +2 pts for Priority_794 Target
    return None, None, 0


# ==============================================================================
# 4. UTILITIES & SCORING ENGINE
# ==============================================================================

def clean_text(raw: str) -> str:
    if not raw:
        return ""
    text = re.sub(r"<[^>]+>", " ", raw)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def fetch_json(url: str, payload: dict = None):
    if payload is not None:
        data_bytes = json.dumps(payload).encode("utf-8")
        headers = {**HEADERS, "Content-Type": "application/json"}
        req = urllib.request.Request(url, data=data_bytes, headers=headers, method="POST")
    else:
        req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=25) as resp:
        return json.loads(resp.read().decode("utf-8"))


def parse_date(raw) -> datetime:
    if not raw:
        return datetime.now(timezone.utc)
    if isinstance(raw, (int, float)):
        ts = raw / 1000.0 if raw > 1e12 else raw
        return datetime.fromtimestamp(ts, tz=timezone.utc)
    raw_str = str(raw).strip()
    try:
        dt = parsedate_to_datetime(raw_str)
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        pass
    try:
        dt = datetime.fromisoformat(raw_str.replace("Z", "+00:00"))
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        return datetime.now(timezone.utc)


def score_job(title: str, description: str, location_hint: str = "", require_location: bool = True, source_name: str = ""):
    title_l = title.lower()
    full_text = f"{source_name} {title} {location_hint} {description}".lower()

    for bad in EXCLUDE_KEYWORDS:
        if bad in title_l:
            return 0, [], [], None

    matched_locs = [loc for loc in TARGET_LOCATIONS if loc in f"{location_hint} {title}".lower()]
    if not matched_locs:
        matched_locs = [loc for loc in TARGET_LOCATIONS if loc in full_text]

    if require_location and not matched_locs:
        return 0, [], [], None

    score = 0
    matched_tags = []

    if matched_locs:
        score += 3

    core_hits = 0
    for kw in CORE_KEYWORDS:
        if kw in full_text:
            score += 3 if kw in title_l else 2
            core_hits += 1
            matched_tags.append(kw)

    # Must match at least one core policy/PA/energy/tech/multilateral domain keyword
    if core_hits == 0:
        return 0, [], [], None

    for sig in SENIORITY_SIGNALS:
        if sig in title_l or f" {sig} " in full_text:
            score += 2
            matched_tags.append(sig)

    db_org, db_pillar, db_boost = match_database_employer(f"{source_name} {title} {description}")
    if db_boost:
        score += db_boost
        matched_tags.append(f"DB:{db_org}")

    return score, sorted(set(matched_locs))[:3], sorted(set(matched_tags))[:7], db_pillar


def make_item(source_name: str, pillar: str, title: str, link: str, loc_str: str, desc: str, pub_raw, score: int, locs: list, tags: list) -> dict:
    loc_display = loc_str or ", ".join(locs) or "EU / Target Hub"
    pillar_tag = f" | {pillar}" if pillar else ""
    return {
        "title": f"[{source_name}{pillar_tag}] {title}",
        "link": link,
        "description": f"Score: {score} | Location: {loc_display} | Tags: {', '.join(tags)}\n\n{desc[:700]}",
        "pub_date": parse_date(pub_raw),
        "score": score,
    }


# ==============================================================================
# 5. ATS & JOB PORTAL SCRAPERS
# ==============================================================================

def fetch_smartrecruiters(board: dict) -> list:
    items = []
    url = f"https://api.smartrecruiters.com/v1/companies/{board['slug']}/postings?limit=100"
    try:
        data = fetch_json(url)
        for job in data.get("content", []):
            title = clean_text(job.get("name", ""))
            job_id = job.get("id", "")
            link = f"https://jobs.smartrecruiters.com/{board['slug']}/{job_id}"
            loc_obj = job.get("location", {}) or {}
            loc_str = ", ".join(filter(None, [loc_obj.get("city"), loc_obj.get("region"), loc_obj.get("country")]))
            dept = (job.get("department") or {}).get("label", "")
            exp = (job.get("experienceLevel") or {}).get("label", "")
            desc = f"{dept} | {exp} | {title}".strip(" |")
            score, locs, tags, db_pillar = score_job(
                title, desc, location_hint=loc_str, require_location=board.get("require_location", True), source_name=board["name"]
            )
            if score >= MIN_SCORE:
                items.append(make_item(board["name"], db_pillar or board.get("pillar", ""), title, link, loc_str, desc, job.get("releasedDate"), score, locs, tags))
    except Exception as exc:
        print(f"[WARN] SmartRecruiters ({board['name']}) failed: {exc}")
    return items


def fetch_greenhouse(board: dict) -> list:
    items = []
    url = f"https://boards-api.greenhouse.io/v1/boards/{board['slug']}/jobs?content=true"
    try:
        data = fetch_json(url)
        for job in data.get("jobs", []):
            title = clean_text(job.get("title", ""))
            link = job.get("absolute_url", "")
            loc_str = (job.get("location") or {}).get("name", "")
            desc = clean_text(job.get("content", ""))
            score, locs, tags, db_pillar = score_job(
                title, desc, location_hint=loc_str, require_location=board.get("require_location", True), source_name=board["name"]
            )
            if score >= MIN_SCORE:
                items.append(make_item(board["name"], db_pillar or board.get("pillar", ""), title, link, loc_str, desc, job.get("updated_at"), score, locs, tags))
    except Exception as exc:
        print(f"[WARN] Greenhouse ({board['name']}) failed: {exc}")
    return items


def fetch_ashby(board: dict) -> list:
    items = []
    url = f"https://api.ashbyhq.com/posting-api/job-board/{board['slug']}"
    try:
        data = fetch_json(url)
        for job in data.get("jobs", []):
            title = clean_text(job.get("title", ""))
            link = job.get("jobUrl", "") or job.get("applyUrl", "")
            loc_str = job.get("location", "")
            dept = job.get("department", "") or job.get("team", "")
            desc = clean_text(job.get("descriptionPlain", "") or job.get("descriptionHtml", ""))
            score, locs, tags, db_pillar = score_job(
                title, f"{dept} {desc}", location_hint=loc_str, require_location=board.get("require_location", True), source_name=board["name"]
            )
            if score >= MIN_SCORE:
                items.append(make_item(board["name"], db_pillar or board.get("pillar", ""), title, link, loc_str, desc, job.get("publishedAt"), score, locs, tags))
    except Exception as exc:
        print(f"[WARN] Ashby ({board['name']}) failed: {exc}")
    return items


def fetch_lever(board: dict) -> list:
    items = []
    url = f"https://api.lever.co/v0/postings/{board['slug']}?mode=json"
    try:
        data = fetch_json(url)
        for job in data if isinstance(data, list) else []:
            title = clean_text(job.get("text", ""))
            link = job.get("hostedUrl", "")
            cats = job.get("categories", {}) or {}
            loc_str = ", ".join(filter(None, [cats.get("location"), cats.get("team"), cats.get("department")]))
            desc = clean_text(job.get("descriptionPlain", "") or job.get("description", ""))
            score, locs, tags, db_pillar = score_job(
                title, desc, location_hint=loc_str, require_location=board.get("require_location", True), source_name=board["name"]
            )
            if score >= MIN_SCORE:
                items.append(make_item(board["name"], db_pillar or board.get("pillar", ""), title, link, loc_str, desc, job.get("createdAt"), score, locs, tags))
    except Exception as exc:
        print(f"[WARN] Lever ({board['name']}) failed: {exc}")
    return items


def fetch_workable(board: dict) -> list:
    items = []
    url = f"https://apply.workable.com/api/v3/accounts/{board['slug']}/jobs"
    try:
        data = fetch_json(url, payload={"query": "", "location": [], "department": [], "worktype": [], "remote": []})
        for job in data.get("results", []):
            title = clean_text(job.get("title", ""))
            shortcode = job.get("shortcode", "")
            link = f"https://apply.workable.com/{board['slug']}/j/{shortcode}/"
            loc_obj = job.get("location", {}) or {}
            loc_str = ", ".join(filter(None, [loc_obj.get("city"), loc_obj.get("country"), "Remote" if job.get("remote") else ""]))
            dept = ", ".join(job.get("department", []) or [])
            desc = f"{dept} | {title} | Energy & climate policy".strip(" |")
            score, locs, tags, db_pillar = score_job(
                title, desc, location_hint=loc_str, require_location=board.get("require_location", True), source_name=board["name"]
            )
            if score >= MIN_SCORE:
                items.append(make_item(board["name"], db_pillar or board.get("pillar", ""), title, link, loc_str, desc, job.get("published"), score, locs, tags))
    except Exception as exc:
        print(f"[WARN] Workable ({board['name']}) failed: {exc}")
    return items


def fetch_jobylon(board: dict) -> list:
    items = []
    try:
        req = urllib.request.Request(board["url"], headers=HEADERS)
        with urllib.request.urlopen(req, timeout=25) as resp:
            page_html = resp.read().decode("utf-8", errors="ignore")
        # Extract Jobylon job links and titles
        matches = re.findall(r'href="(https://emp\.jobylon\.com/jobs/[^"]+)"[^>]*>(.*?)</a>', page_html, re.S)
        seen_links = set()
        for link, raw_title in matches:
            title = clean_text(raw_title)
            if not title or len(title) < 5 or link in seen_links or "apply" == title.lower():
                continue
            seen_links.add(link)
            desc = f"Public Affairs Consultancy role at {board['name']}: {title}"
            score, locs, tags, db_pillar = score_job(
                title, desc, location_hint=title, require_location=board.get("require_location", True), source_name=board["name"]
            )
            if score >= MIN_SCORE:
                items.append(make_item(board["name"], db_pillar or board.get("pillar", ""), title, link, "", desc, None, score, locs, tags))
    except Exception as exc:
        print(f"[WARN] Jobylon ({board['name']}) failed: {exc}")
    return items


def fetch_rss_source(source: dict) -> list:
    items = []
    try:
        req = urllib.request.Request(source["url"], headers=HEADERS)
        with urllib.request.urlopen(req, timeout=20) as resp:
            xml_bytes = resp.read()
        root = ET.fromstring(xml_bytes)

        entries = root.findall(".//item")
        is_atom = False
        if not entries:
            entries = root.findall(".//{http://www.w3.org/2005/Atom}entry")
            is_atom = True

        for el in entries:
            if is_atom:
                ns = "{http://www.w3.org/2005/Atom}"
                title = clean_text(el.findtext(f"{ns}title", ""))
                link_el = el.find(f"{ns}link")
                link = link_el.attrib.get("href", "") if link_el is not None else ""
                desc = clean_text(el.findtext(f"{ns}summary", "") or el.findtext(f"{ns}content", ""))
                pub_raw = el.findtext(f"{ns}updated", "") or el.findtext(f"{ns}published", "")
            else:
                title = clean_text(el.findtext("title", ""))
                link = clean_text(el.findtext("link", ""))
                desc = clean_text(el.findtext("description", ""))
                pub_raw = el.findtext("pubDate", "")

            if not title or not link:
                continue

            score, locs, tags, db_pillar = score_job(
                title, desc, require_location=source.get("require_location", True), source_name=source["name"]
            )
            if score >= MIN_SCORE:
                items.append(make_item(source["name"], db_pillar or source.get("pillar", ""), title, link, "", desc, pub_raw, score, locs, tags))
    except Exception as exc:
        print(f"[WARN] RSS ({source['name']}) failed: {exc}")
    return items


def fetch_reliefweb_jobs() -> list:
    items = []
    try:
        data = fetch_json(RELIEFWEB_API_URL)
        for entry in data.get("data", []):
            fields = entry.get("fields", {})
            title = clean_text(fields.get("title", ""))
            link = fields.get("url", "")
            body = clean_text(fields.get("body", ""))
            countries = [c.get("name", "") for c in fields.get("country", [])]
            cities = [c.get("name", "") for c in fields.get("city", [])]
            orgs = [s.get("name", "") for s in fields.get("source", [])]
            loc_str = ", ".join(cities + countries)
            org_str = ", ".join(orgs) or "ReliefWeb"

            score, locs, tags, db_pillar = score_job(
                title, body, location_hint=loc_str, require_location=True, source_name=org_str
            )
            if score >= MIN_SCORE:
                items.append(make_item(f"ReliefWeb | {org_str}", db_pillar or "Multilateral / NGO", title, link, loc_str, body, fields.get("date", {}).get("created", ""), score, locs, tags))
    except Exception as exc:
        print(f"[WARN] ReliefWeb API failed: {exc}")
    return items


# ==============================================================================
# 6. BUILD RSS 2.0 XML
# ==============================================================================

def build_rss(items: list, output_path: str = "feed.xml"):
    seen = set()
    unique_items = []
    for item in sorted(items, key=lambda x: (x["pub_date"], x["score"]), reverse=True):
        if item["link"] not in seen:
            seen.add(item["link"])
            unique_items.append(item)

    rss = ET.Element("rss", version="2.0")
    channel = ET.SubElement(rss, "channel")
    ET.SubElement(channel, "title").text = "Aron Miszlivetz — Master Database EU Policy, Energy & Tech Jobs"
    ET.SubElement(channel, "link").text = "https://github.com/"
    ET.SubElement(channel, "description").text = (
        "Daily auto-updated RSS feed built from the 1,200-organisation Master Career Database: "
        "EU policy, public affairs consultancies, government affairs, energy, tech/digital/AI & multilateral roles."
    )
    ET.SubElement(channel, "lastBuildDate").text = format_datetime(datetime.now(timezone.utc))

    for item in unique_items[:200]:
        it = ET.SubElement(channel, "item")
        ET.SubElement(it, "title").text = item["title"]
        ET.SubElement(it, "link").text = item["link"]
        ET.SubElement(it, "guid").text = item["link"]
        ET.SubElement(it, "pubDate").text = format_datetime(item["pub_date"])
        ET.SubElement(it, "description").text = item["description"]

    tree = ET.ElementTree(rss)
    ET.indent(tree, space="  ")
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    print(f"Wrote {len(unique_items[:200])} matched jobs to {output_path}")


if __name__ == "__main__":
    all_jobs = []
    for b in SMARTRECRUITERS_BOARDS:
        all_jobs.extend(fetch_smartrecruiters(b))
    for b in GREENHOUSE_BOARDS:
        all_jobs.extend(fetch_greenhouse(b))
    for b in ASHBY_BOARDS:
        all_jobs.extend(fetch_ashby(b))
    for b in LEVER_BOARDS:
        all_jobs.extend(fetch_lever(b))
    for b in WORKABLE_BOARDS:
        all_jobs.extend(fetch_workable(b))
    for b in JOBYLON_BOARDS:
        all_jobs.extend(fetch_jobylon(b))
    for src in RSS_SOURCES:
        all_jobs.extend(fetch_rss_source(src))
    all_jobs.extend(fetch_reliefweb_jobs())
    build_rss(all_jobs, "feed.xml")
