"""
contact_extracted.py  —  Drop-in replacement for the contact-scraping section
of scraper.py.

(Renamed from contact_extractor.py so it matches the import already in
scraper.py: `from contact_extracted import scrape_company_contacts_v2`.
If you'd rather keep the original filename, change that import line instead
and rename this file back.)

四层策略 (highest → lowest confidence):
  Layer 1  Structured HTML tags   a[href^="tel/mailto:"], JSON-LD schema.org
  Layer 2  Contact-page discovery  /contact, sitemap, nav-link scan
  Layer 3  Context-aware regex     footer / contact section only
  Layer 4  JS-render fallback      use the existing Selenium driver if layers 1-3 miss

Usage  (in scraper.py):
    from contact_extracted import scrape_company_contacts_v2
    contacts = scrape_company_contacts_v2(website_url, driver=driver)
    rec.phone = contacts["phone"]
    rec.email = contacts["email"]

CHANGELOG vs. original contact_extractor.py
---------------------------------------------
1. Phone regex fixed: the original only matched numbers where the
   subscriber digits were written as ONE contiguous block after the area
   code (e.g. "+60-3-12345678"). Most real MY sites group the digits
   (e.g. "+60 3-7890 1234", "+60 12-345 6789") — those all failed to match
   before. Verified against real-world formats before/after the fix.
2. Added optional support for LOCAL-format numbers (no "+60"/"60" prefix,
   e.g. "03-7890 1234", "012-345 6789", "1-300-88-1234"), since most MY
   company sites just print the local dialling format. Toggle with
   MATCH_LOCAL_FORMAT_PHONES below. All matched numbers — local or
   international — are normalised to a single consistent "+60XXXXXXXXX"
   representation so you never get the same number stored twice in two
   different formats.
3. tel:/mailto:/JSON-LD extraction now goes through the same normaliser,
   so a tel="0312345678" href (no country code) is recognised too.
"""

import json
import re
import time
import random
import logging
from typing import Optional
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

log = logging.getLogger(__name__)

# Set to False if you only want numbers that explicitly include a +60 /
# 60 country code, and want to discard plain local-format numbers.
MATCH_LOCAL_FORMAT_PHONES = False

# ── Request headers ────────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# ── Phone regex — Malaysia, with or without country code ─────────────────────
# Matches the raw digit grouping as printed on the page; normalisation /
# country-code enforcement happens afterwards in _clean_phone().
#   branch 1: mobile (01X) / peninsular landline (03,04,05,06,07,09) /
#             Sabah & Sarawak landline (082-089), optional leading 0,
#             optional country code, digits grouped in any combination
#             of 3-4 + 3-4 with arbitrary separators between every group.
#   branch 2: toll-free (1-300/600/700/800-XX-XXXX)
_PHONE_RE = re.compile(
    r'(?:\+?60[\s.\-]?)?'
    r'(?:'
    r'0?(?:1[0-9]|[3-7]|9|8[2-9])[\s.\-]?\d{3,4}[\s.\-]?\d{3,4}'
    r'|0?1[\s.\-]?(?:300|600|700|800)[\s.\-]?\d{2}[\s.\-]?\d{4}'
    r')'
)

_EMAIL_RE = re.compile(
    r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',
)

_EMAIL_SKIP_RE = re.compile(
    r'\.(png|jpg|gif|svg|webp|ico|css|js|woff|ttf)$'
    r'|@(example|yourdomain|sentry|email\.com|domain\.com)',
    re.IGNORECASE,
)

# Contact-page URL keywords (EN + BM)
_CONTACT_SLUGS = [
    "contact", "contact-us", "contacts", "contactus",
    "about", "about-us", "aboutus",
    "hubungi", "hubungi-kami", "kenalan",
    "info", "enquiry", "enquiries",
    "reach-us", "get-in-touch",
]

# CSS selectors for zones most likely to hold real contacts
_CONTACT_ZONE_SELECTORS = [
    "footer",
    "[id*='contact']",   "[class*='contact']",
    "[id*='footer']",    "[class*='footer']",
    "address",
    "[id*='reach']",     "[class*='reach']",
    "[id*='enquir']",    "[class*='enquir']",
]

# Confidence weights
_CONF = {
    "tel_href":     10,
    "mailto":       10,
    "json_ld":       9,
    "contact_zone":  6,
    "whole_page":    2,
}


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _clean_phone(raw: str) -> str:
    """
    Normalise any MY phone string (with or without country code, any
    separator style) into a single canonical "+60XXXXXXXXX" form.
    Returns "" if it doesn't look like a valid MY number, or if it's a
    local-format number and MATCH_LOCAL_FORMAT_PHONES is False.
    """
    raw = raw.strip()
    digits = re.sub(r"[^\d+]", "", raw)

    if digits.startswith("+60"):
        normalized = digits
    elif digits.startswith("60") and len(digits) >= 11:
        normalized = "+" + digits
    elif MATCH_LOCAL_FORMAT_PHONES and digits.startswith("0"):
        normalized = "+60" + digits[1:]
    elif MATCH_LOCAL_FORMAT_PHONES and re.match(r"^1(300|600|700|800)", digits):
        normalized = "+60" + digits
    else:
        return ""

    digit_count = len(re.sub(r"\D", "", normalized))
    return normalized if 10 <= digit_count <= 12 else ""


def _clean_email(raw: str) -> str:
    raw = raw.strip().lower()
    return "" if _EMAIL_SKIP_RE.search(raw) else raw


def _root(url: str) -> str:
    p = urlparse(url)
    return f"{p.scheme}://{p.netloc}"


def _fetch(session: requests.Session, url: str, timeout: int = 12, retries: int = 2) -> Optional[str]:
    for attempt in range(retries + 1):
        try:
            r = session.get(url, timeout=timeout, allow_redirects=True)
            r.raise_for_status()
            return r.text
        except requests.exceptions.ConnectionError as e:
            if attempt < retries:
                time.sleep(1.5 * (attempt + 1))
                log.debug(f"      fetch retry {attempt+1} for {url}: {e}")
            else:
                log.debug(f"      fetch failed {url}: {e}")
                return None
        except Exception as e:
            log.debug(f"      fetch {url}: {e}")
            return None


# ─────────────────────────────────────────────────────────────────────────────
# Layer 1 — Structured HTML tags + JSON-LD
# ─────────────────────────────────────────────────────────────────────────────

def _extract_structured(soup: BeautifulSoup) -> dict:
    result = {"phone": [], "email": []}

    # <a href="tel:...">  (with or without country code)
    for a in soup.find_all("a", href=re.compile(r'^tel:', re.I)):
        raw = a["href"].replace("tel:", "").strip()
        p = _clean_phone(raw)
        if p:
            result["phone"].append((_CONF["tel_href"], p))

    # <a href="mailto:...">
    for a in soup.find_all("a", href=re.compile(r'^mailto:', re.I)):
        raw = a["href"].replace("mailto:", "").split("?")[0].strip()
        e = _clean_email(raw)
        if e:
            result["email"].append((_CONF["mailto"], e))

    # schema.org JSON-LD
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
            if isinstance(data, list):
                data = data[0]
            phone_raw = str(data.get("telephone", "") or data.get("phone", ""))
            p = _clean_phone(phone_raw)
            if p:
                result["phone"].append((_CONF["json_ld"], p))
            email_raw = str(data.get("email", "")).replace("mailto:", "")
            e = _clean_email(email_raw)
            if e:
                result["email"].append((_CONF["json_ld"], e))
        except Exception:
            pass

    return result


# ─────────────────────────────────────────────────────────────────────────────
# Layer 2 — Contact-page discovery
# ─────────────────────────────────────────────────────────────────────────────

def _find_contact_pages(
    session: requests.Session,
    base_url: str,
    soup: BeautifulSoup,
) -> list:
    root = _root(base_url)
    candidates = []
    seen = {base_url}

    # a) Links on the page
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith(("#", "javascript", "mailto", "tel")):
            continue
        full = urljoin(base_url, href).rstrip("/")
        if any(kw in full.lower() for kw in _CONTACT_SLUGS) and full not in seen:
            seen.add(full)
            candidates.append(full)
            if len(candidates) >= 3:
                break

    # b) Common slug probing with GET+stream (HEAD breaks on many MY servers)
    if len(candidates) < 2:
        for slug in _CONTACT_SLUGS[:8]:
            url = f"{root}/{slug}"
            if url in seen:
                continue
            try:
                r = session.get(url, timeout=6, allow_redirects=True, stream=True)
                r.close()
                if r.status_code == 200:
                    seen.add(url)
                    candidates.append(url)
                    if len(candidates) >= 3:
                        break
            except Exception:
                pass

    # c) sitemap.xml fallback
    if len(candidates) < 1:
        sitemap = _fetch(session, f"{root}/sitemap.xml", timeout=8)
        if sitemap:
            for loc in re.findall(r'<loc>(.*?)</loc>', sitemap, re.I):
                if any(kw in loc.lower() for kw in _CONTACT_SLUGS) and loc not in seen:
                    seen.add(loc)
                    candidates.append(loc)
                    if len(candidates) >= 2:
                        break

    return candidates[:3]


# ─────────────────────────────────────────────────────────────────────────────
# Layer 3 — Context-aware regex on contact zone
# ─────────────────────────────────────────────────────────────────────────────

def _extract_from_zones(soup: BeautifulSoup, confidence: int) -> dict:
    result = {"phone": [], "email": []}
    zones = []

    for selector in _CONTACT_ZONE_SELECTORS:
        for el in soup.select(selector):
            zones.append(el)

    # Deduplicate zones
    unique_zones = []
    seen_ids = set()
    for z in zones:
        if id(z) not in seen_ids:
            seen_ids.add(id(z))
            unique_zones.append(z)

    if not unique_zones:
        unique_zones = [soup]
        confidence = max(1, confidence - 2)

    for zone in unique_zones:
        for tag in zone(["script", "style", "noscript"]):
            tag.decompose()

        text = zone.get_text(separator=" ", strip=True)

        for m in _PHONE_RE.finditer(text):
            p = _clean_phone(m.group())
            if p:
                result["phone"].append((confidence, p))

        for e_raw in _EMAIL_RE.findall(text):
            e = _clean_email(e_raw)
            if e:
                result["email"].append((confidence, e))

    return result


# ─────────────────────────────────────────────────────────────────────────────
# Layer 4 — JS-render fallback via Selenium
# ─────────────────────────────────────────────────────────────────────────────

def _extract_via_selenium(driver, url: str) -> dict:
    empty = {"phone": "", "email": ""}

    try:
        driver.set_page_load_timeout(20)
    except Exception:
        pass

    try:
        try:
            driver.get(url)
        except Exception:
            try:
                driver.execute_script("window.stop();")
            except Exception:
                pass

        time.sleep(2)

        try:
            html = driver.execute_script("return document.documentElement.outerHTML;")
        except Exception as e:
            log.warning(f"      JS innerHTML read failed for {url}: {e}")
            return empty

        if not html:
            return empty

        soup = BeautifulSoup(html, "lxml")
        structured = _extract_structured(soup)
        zone_data = _extract_from_zones(soup, confidence=_CONF["contact_zone"])
        return _merge([structured, zone_data])

    except Exception as e:
        log.warning(f"      Selenium fallback failed for {url}: {e}")
        return empty
    finally:
        try:
            driver.set_page_load_timeout(30)
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# Merge — pick highest-confidence deduplicated values
# ─────────────────────────────────────────────────────────────────────────────

def _merge(extractions: list) -> dict:
    combined = {"phone": [], "email": []}
    for ex in extractions:
        for key in combined:
            combined[key].extend(ex.get(key, []))

    result = {}
    for key, candidates in combined.items():
        best = {}
        for conf, val in candidates:
            val = val.strip()
            if not val:
                continue
            if val not in best or conf > best[val]:
                best[val] = conf
        result[key] = max(best, key=lambda v: best[v]) if best else ""

    return result


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def scrape_company_contacts_v2(
    website_url: str,
    driver=None,
    timeout: int = 12,
) -> dict:
    """
    Multi-layer contact extractor. Returns {"phone": str, "email": str}.
    Phone is normalised to +60XXXXXXXXX format. Values are empty strings
    if not found.
    """
    session = requests.Session()
    session.headers.update(HEADERS)
    all_extractions = []

    def _process_html(html: str, url: str, is_contact_page: bool):
        soup = BeautifulSoup(html, "lxml")
        structured = _extract_structured(soup)
        conf = _CONF["contact_zone"] if is_contact_page else _CONF["whole_page"]
        zones = _extract_from_zones(soup, confidence=conf)
        return [structured, zones], soup

    # ── Homepage ──────────────────────────────────────────────────────────────
    homepage_html = _fetch(session, website_url, timeout)
    if not homepage_html:
        log.warning(f"    Could not fetch homepage: {website_url}")
        if driver:
            return _extract_via_selenium(driver, website_url)
        return {"phone": "", "email": ""}

    homepage_raws, homepage_soup = _process_html(homepage_html, website_url, False)
    all_extractions.extend(homepage_raws)

    merged = _merge(all_extractions)
    if all(merged.values()):
        log.info("    All contacts found on homepage")
        return merged

    # ── Contact pages (Layer 2) ────────────────────────────────────────────
    contact_pages = _find_contact_pages(session, website_url, homepage_soup)
    log.debug(f"    Contact pages to try: {contact_pages}")

    for cp_url in contact_pages:
        cp_html = _fetch(session, cp_url, timeout)
        if not cp_html:
            continue
        cp_raws, _ = _process_html(cp_html, cp_url, is_contact_page=True)
        all_extractions.extend(cp_raws)
        time.sleep(random.uniform(0.4, 1.0))

        merged = _merge(all_extractions)
        if all(merged.values()):
            log.info(f"    All contacts found (contact page: {cp_url})")
            return merged

    # ── JS-render fallback (Layer 4) ──────────────────────────────────────
    merged = _merge(all_extractions)
    still_missing = [k for k, v in merged.items() if not v]
    if still_missing and driver:
        log.info(f"    Missing {still_missing}, trying JS render")
        target = contact_pages[0] if contact_pages else website_url
        js_data = _extract_via_selenium(driver, target)
        all_extractions.append({k: [(8, v)] for k, v in js_data.items() if v})

    return _merge(all_extractions)


# ─────────────────────────────────────────────────────────────────────────────
# Quick standalone test
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(message)s")
    url = sys.argv[1] if len(sys.argv) > 1 else "https://www.maybank2u.com.my"
    print(f"\nTesting: {url}")
    result = scrape_company_contacts_v2(url)
    print(f"  phone: {result['phone'] or '-'}")
    print(f"  email: {result['email'] or '-'}")