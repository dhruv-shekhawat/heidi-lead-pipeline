"""Company-website helpers (scripts/websites.py) shared by light_enrich.py and extract_contacts.py.

All requests go through common.fetch (robots.txt, cache, 403/429 -> Blocked).
"""
import re
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup

import common

TIMEOUT = 10
DELAY = 1.0


def homepage(domain):
    """Fetch a company homepage. Returns (entry, None) or (None, reason)."""
    d = common.norm_domain(domain)
    if not d:
        return None, "keine Domain"
    cands = [f"https://{d}/", f"https://www.{d}/", f"http://www.{d}/", f"http://{d}/"]
    last = "unbekannt"
    for u in dict.fromkeys(cands):
        try:
            return common.fetch(u, delay=DELAY, timeout=TIMEOUT), None
        except common.Blocked as e:
            return None, str(e)  # never retry a host that blocked us
        except common.FetchError as e:
            last = str(e)
            if "robots" in last:
                return None, last
    return None, last


def page_url(entry):
    return entry.get("final_url") or entry["url"]


def links(entry):
    soup = BeautifulSoup(entry.get("text", ""), "lxml")
    out = []
    for a in soup.select("a[href]"):
        href = a["href"].strip()
        if href.startswith(("javascript:", "#")):
            continue
        out.append((a.get_text(" ", strip=True), urljoin(page_url(entry), href)))
    return out


def same_site(url, entry):
    h1 = common.norm_domain(url)
    h2 = common.norm_domain(page_url(entry))
    return h1 == h2 or h1.endswith("." + h2) or h2.endswith("." + h1)


def find_link(entry, pattern):
    rx = re.compile(pattern, re.I)
    for text, url in links(entry):
        if url.lower().startswith(("http://", "https://")) and same_site(url, entry) \
                and (rx.search(text) or rx.search(urlparse(url).path)):
            return url
    return None


def subpage(home, pattern, fallback_path):
    """Find + fetch a linked subpage (Impressum, Kontakt). Returns (entry, url, reason)."""
    url = find_link(home, pattern)
    tried = []
    for u in ([url] if url else []) + [urljoin(page_url(home), fallback_path)]:
        if u in tried:
            continue
        tried.append(u)
        try:
            return common.fetch(u, delay=DELAY, timeout=TIMEOUT), u, None
        except common.Blocked as e:
            return None, u, str(e)
        except common.FetchError as e:
            reason = str(e)
    return None, None, f"nicht gefunden ({reason})" if tried else "nicht gefunden"


def impressum(home):
    return subpage(home, r"impressum|imprint", "/impressum")


def kontakt(home):
    return subpage(home, r"kontakt|contact", "/kontakt")
