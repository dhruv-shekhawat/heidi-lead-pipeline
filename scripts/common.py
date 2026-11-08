"""Shared helpers: cached + robots-aware fetching, normalisation, PLZ parsing.

Cache: data/cache/<sha1>.json, one file per request. Key = sha1 of the URL
(for POST requests the sorted form body is appended to the URL before hashing).
Each entry stores url, final_url, status, content_type, text, error, fetched.
Successful pages AND definitive failures (HTTP 4xx, robots disallow) are cached, so reruns and
offline dry-runs reproduce the same result without network access. Transient failures (timeouts,
HTTP 5xx) are cached with "transient": true – online runs retry them, offline runs replay them.
Set HEIDI_OFFLINE=1 to forbid network access entirely (cache misses -> FetchError).
"""
import hashlib
import json
import os
import re
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode, urljoin, urlparse
from urllib.robotparser import RobotFileParser

import requests

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
CACHE = DATA / "cache"
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")
OFFLINE = os.environ.get("HEIDI_OFFLINE") == "1"
# Hinweis marker for an agent-written Einheiten figure >= the Enterprise threshold: score.py ignores the figure
# until it is verified (manual_fixes setting Einheiten removes the marker)
AGENT_EINHEITEN_MARK = "Agent-Zahl, prüfen"
AGENT_MARK_RX = re.compile(r"\s*(?:\|\s*)?Agent-Zahl, prüfen(?: \([^)]*\))?")

session = requests.Session()
session.headers.update({
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.5",
})


class Blocked(Exception):
    """HTTP 403 or 429."""


class FetchError(Exception):
    """Anything else: robots disallow, timeout, non-200, offline cache miss."""


# ------------------------------------------------------------------ cache
def cache_key(url, data=None):
    k = url if not data else url + "\n" + urlencode(sorted(data.items()))
    return hashlib.sha1(k.encode("utf-8")).hexdigest()


def cache_path(url, data=None):
    return CACHE / f"{cache_key(url, data)}.json"


def cache_read(url, data=None):
    p = cache_path(url, data)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except ValueError:
            return None
    return None


def cache_write(url, data, entry):
    CACHE.mkdir(parents=True, exist_ok=True)
    cache_path(url, data).write_text(json.dumps(entry, ensure_ascii=False), encoding="utf-8")


def _raise_for(entry):
    if entry.get("error"):
        err = entry["error"]
        if err.startswith("HTTP 403") or err.startswith("HTTP 429"):
            raise Blocked(err)
        raise FetchError(err)
    return entry


# ------------------------------------------------------------------ robots
_robots = {}
_last_hit = {}


def robots_for(url):
    """RobotFileParser for the URL's host (cached in memory and on disk)."""
    p = urlparse(url)
    base = f"{p.scheme}://{p.netloc}"
    if base in _robots:
        return _robots[base]
    rurl = base + "/robots.txt"
    rp = RobotFileParser()
    entry = cache_read(rurl)
    if entry is None and not OFFLINE:
        try:
            r = session.get(rurl, timeout=10, allow_redirects=True)
            entry = {"url": rurl, "status": r.status_code,
                     "text": r.text if r.status_code == 200 else "",
                     "fetched": datetime.now().isoformat(timespec="seconds")}
        except requests.RequestException as e:
            entry = {"url": rurl, "status": 0, "text": "", "error": f"robots: {type(e).__name__}"}
        cache_write(rurl, None, entry)
    if entry is None:  # offline and never fetched
        rp = None
    elif entry.get("status") in (401, 403):
        rp.disallow_all = True
    elif entry.get("status") == 200:
        rp.parse(entry["text"].splitlines())
    else:  # 404, other 4xx/5xx, network error: no usable robots.txt -> allow
        rp.allow_all = True
    _robots[base] = rp
    return rp


def robots_allowed(url):
    rp = robots_for(url)
    if rp is None:
        raise FetchError("offline: robots.txt nicht im Cache")
    return rp.can_fetch(UA, url)


def crawl_delay(url):
    rp = robots_for(url)
    try:
        d = rp.crawl_delay(UA) if rp else None
    except Exception:
        d = None
    return float(d) if d else 0.0


# ------------------------------------------------------------------ fetch
# soft hyphen + zero-width characters (and their HTML entities) split words like "Miet\u00adverwaltung"
INVISIBLE_RX = re.compile(r"[\u00ad\u200b\u200c\u200d\u2060\ufeff]|&(?:shy|zwnj|zwj);|"
                          r"&#(?:173|8203|8204|8205|8288|65279);|&#x(?:ad|200b|200c|200d|2060|feff);", re.I)


def clean_text(s):
    """Remove soft hyphens and zero-width characters (raw or as entities) before any parsing."""
    return INVISIBLE_RX.sub("", s) if s else s


def fetch(url, **kw):
    """Return cache entry dict (url, final_url, status, content_type, text); text is cleaned of
    soft hyphens / zero-width characters (the cache file keeps the original bytes).

    Raises Blocked on 403/429, FetchError otherwise. Uses cache first.
    """
    entry = _fetch(url, **kw)
    if entry.get("text"):
        entry = dict(entry, text=clean_text(entry["text"]))
    return entry


def _fetch(url, *, method="GET", data=None, headers=None, delay=1.0, timeout=20):
    entry = cache_read(url, data)
    if entry is not None and not (entry.get("transient") and not OFFLINE):
        entry["cached"] = True  # transient failures are only replayed offline; online they are retried
        return _raise_for(entry)
    if OFFLINE:
        raise FetchError("offline: nicht im Cache")
    if not robots_allowed(url):
        entry = {"url": url, "status": 0, "error": "robots.txt verbietet Zugriff"}
        cache_write(url, data, entry)
        return _raise_for(entry)
    host = urlparse(url).netloc
    wait = max(delay, crawl_delay(url)) - (time.time() - _last_hit.get(host, 0))
    if wait > 0:
        time.sleep(wait)
    _last_hit[host] = time.time()
    try:
        r = session.request(method, url, data=data, headers=headers, timeout=timeout,
                            allow_redirects=True)
    except requests.Timeout:
        cache_write(url, data, {"url": url, "status": 0, "error": "Timeout", "transient": True})
        raise FetchError("Timeout")  # transient: retried online, replayed offline
    except requests.exceptions.SSLError:
        entry = {"url": url, "status": 0, "error": "SSL-Fehler"}
        cache_write(url, data, entry)
        return _raise_for(entry)
    except requests.ConnectionError:
        entry = {"url": url, "status": 0, "error": "Verbindung fehlgeschlagen"}
        cache_write(url, data, entry)
        return _raise_for(entry)
    except requests.RequestException as e:
        raise FetchError(type(e).__name__)
    entry = {"url": url, "final_url": r.url, "status": r.status_code,
             "content_type": r.headers.get("content-type", ""),
             "fetched": datetime.now().isoformat(timespec="seconds")}
    if r.status_code != 200:
        entry["error"] = f"HTTP {r.status_code}"
    elif not re.search(r"html|json|xml|text", entry["content_type"], re.I):
        entry["error"] = f"kein HTML ({entry['content_type'][:40]})"
    else:
        if not r.encoding or r.encoding.lower() == "iso-8859-1":
            r.encoding = r.apparent_encoding
        entry["text"] = r.text
    if r.status_code >= 500:  # transient: retried online, replayed offline (429 is cached for good)
        entry["transient"] = True
        cache_write(url, data, entry)
        return _raise_for(entry)
    cache_write(url, data, entry)
    entry["cached"] = False
    return _raise_for(entry)


def html_lines(html):
    """Visible text of an HTML page as a list of non-empty lines (line breaks kept).
    For the line parsers (Impressum name, contact person): head/script/style/noscript/svg/template dropped."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(clean_text(html or ""), "lxml")
    for t in soup(["head", "script", "style", "noscript", "svg", "template"]):
        t.decompose()
    for br in soup.find_all("br"):
        br.replace_with("\n")
    text = soup.get_text("\n")
    return [ln for ln in (re.sub(r"[ \t\xa0]+", " ", x).strip() for x in text.split("\n")) if ln]


TEXT_ATTRS = ("alt", "title", "aria-label")


def _ld_strings(raw):
    """Prose string values of a JSON-LD block (description, slogan, …), no JSON syntax or URLs."""
    try:
        stack, out = [json.loads(raw)], []
    except ValueError:
        out = []
        for m in re.findall(r'"(?:description|slogan|text|headline)"\s*:\s*"((?:[^"\\]|\\.)*)"', raw):
            try:
                out.append(json.loads(f'"{m}"'))
            except ValueError:
                out.append(m)
        return out
    while stack:
        x = stack.pop()
        if isinstance(x, dict):
            stack.extend(reversed(list(x.values())))  # pop() then yields document order
        elif isinstance(x, list):
            stack.extend(reversed(x))
        elif isinstance(x, str) and " " in x.strip() and not re.match(r"https?:", x):
            out.append(x)
    return out


def page_text(html):
    """Full text of a page for keyword detection: everything html_lines keeps PLUS text that is
    not rendered or only shown later – <template>/<noscript> content (slider slides, lazy
    sections), JSON-LD descriptions, alt/title/aria-label and plain-text data-* attributes
    (slider captions, counter targets such as data-to="1008")."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(clean_text(html or ""), "lxml")
    extra = [v for t in soup.select('script[type="application/ld+json"]') for v in _ld_strings(t.get_text())]
    for t in soup(["head", "script", "style", "svg"]):
        t.decompose()
    for tpl in soup("template"):  # bs4 excludes template strings from get_text(): re-parse them
        extra.append(BeautifulSoup(tpl.decode_contents(), "lxml").get_text("\n"))
        tpl.decompose()
    for ns in soup("noscript"):  # noscript content is parsed as text: re-parse it
        extra.append(BeautifulSoup(ns.get_text(" "), "lxml").get_text("\n"))
        ns.decompose()
    for el in soup.find_all(True):
        for k, v in el.attrs.items():
            if (k in TEXT_ATTRS or (k.startswith("data-") and k not in ("data-src", "data-srcset", "data-href",
                                                                         "data-link", "data-url"))) \
                    and isinstance(v, str) and 2 <= len(v) <= 300 and re.search(r"[A-Za-zÄÖÜäöü0-9]", v) \
                    and not re.search(r"[{}<>]|https?:|/|\.(png|jpe?g|svg|webp)\b", v):
                extra.append(v)
    for br in soup.find_all("br"):
        br.replace_with("\n")
    lines = soup.get_text("\n").split("\n") + extra
    return "\n".join(ln for ln in (re.sub(r"[ \t\xa0]+", " ", x).strip() for x in lines) if ln)


# ------------------------------------------------------------ normalisation
LEGAL = r"\b(gmbh|mbh|ug|haftungsbeschraenkt|ag|kg|ohg|e\s?k|gbr|co|eg|ev|partg|se|inh|inhaber)\b"
NOT_COMPANY_DOMAINS = (
    "facebook.", "instagram.", "linkedin.", "xing.", "twitter.", "x.com", "youtube.",
    "immobilienscout24.", "immowelt.", "gelbeseiten.", "11880.", "google.", "goo.gl",
    "jacasa.", "wix.com", "jimdo", "t-online.de", "gmx.", "web.de", "yelp.", "golocal.",
    "dasoertliche.", "meinestadt.", "branchenbuch", "cylex", "werkenntdenbesten.",
)


def norm_tokens(s):
    """Lowercase, umlauts -> ae/oe/ue/ss, Rechtsform tokens and branch suffixes removed."""
    s = (s or "").lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        s = s.replace(a, b)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.replace("&", " ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(LEGAL, " ", s)
    s = re.sub(r"\b(zweig)?(nl|niederlassung|filiale|standort)\b", " ", s)
    toks = s.split()
    while len(toks) > 1 and toks[-1] == "berlin":  # branch suffix "… Berlin"
        toks.pop()
    return toks


def norm_name(s):
    """Matching key: tokens joined without spaces/hyphens/punctuation ("abc-hv" == "ABC HV")."""
    return "".join(norm_tokens(s))


def norm_domain(url):
    url = (url or "").strip().lower()
    if not url:
        return ""
    if "://" not in url:
        url = "http://" + url
    host = urlparse(url).netloc.split("@")[-1].split(":")[0]
    return host[4:] if host.startswith("www.") else host


def is_company_domain(d):
    return bool(d) and not any(x in d for x in NOT_COMPANY_DOMAINS)


PLZ_RE = re.compile(r"\b1\d{4}\b")


def plz_of(v):
    m = PLZ_RE.search(str(v or ""))
    return m.group(0) if m else ""


def abs_url(base, href):
    return urljoin(base, href)
