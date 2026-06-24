import re
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote_plus, urlparse, parse_qs, unquote
from config import BLACKLIST_DOMAINS, BAD_DOMAINS, PRIORITY_DOMAINS, HEADERS


def clean_product_name(name: str) -> str:
    s = str(name or "").strip()
    s = re.sub(r"\([^)]*(цвет|red|black|blue|green|white|желт|красн|черн|син|бел|202[0-9])[^)]*\)", " ", s, flags=re.I)
    s = re.sub(r"\b(новый|new|202[0-9]|год|цвет|красный|черный|чёрный|синий|белый|желтый|зелёный|зеленый)\b", " ", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip() or str(name or "").strip()


def is_blacklisted(url: str) -> bool:
    u = (url or "").lower()
    return any(domain in u for domain in BLACKLIST_DOMAINS)


def is_bad_url(url: str) -> bool:
    u = (url or "").lower()
    return is_blacklisted(u) or any(x in u for x in BAD_DOMAINS)


def clean_ddg_url(url: str) -> str:
    if "duckduckgo.com/l/" in (url or ""):
        qs = parse_qs(urlparse(url).query)
        if "uddg" in qs:
            return unquote(qs["uddg"][0])
    return url or ""


def score_url(url: str, product_name: str) -> int:
    u = (url or "").lower()
    p = clean_product_name(product_name).lower()
    if is_bad_url(u):
        return -1000
    score = 0
    if any(d in u for d in PRIORITY_DOMAINS):
        score += 100
    for token in re.findall(r"[a-zA-Zа-яА-Я0-9]+", p):
        if len(token) > 1 and token.lower() in u:
            score += 8
    if any(x in u for x in ["product", "catalog", "character", "spec", "harakter", "kharakteristiki", "products"]):
        score += 20
    return score


def build_queries(product_name: str, category: str) -> list[str]:
    clean = clean_product_name(product_name)
    queries = [
        f'"{clean}" характеристики',
        f'"{clean}" технические характеристики',
        f'"{clean}" specs',
        f'{clean} {category} характеристики',
        f'{clean} купить характеристики',
        f'{clean} инструкция характеристики',
        f'{clean} паспорт характеристики',
    ]
    priority_sites = [
        "rollingmoto.ru", "motomarine.ru", "vodnik.ru", "mymotors.ru",
        "hondaset.ru", "rus-lodki.ru", "tulin-lodki.ru", "lodki-piter.ru",
    ]
    for site in priority_sites:
        queries.append(f'"{clean}" site:{site}')
    queries += [
        f'{clean} характеристики',
        f'{clean} технические характеристики',
    ]
    return list(dict.fromkeys(queries))


def search_serper(query: str, api_key: str, max_results: int = 10):
    if not api_key:
        return [], [], "SERPER_API_KEY не найден"
    try:
        r = requests.post(
            "https://google.serper.dev/search",
            json={"q": query, "gl": "ru", "hl": "ru", "num": max_results},
            headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
            timeout=20,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        return [], [], f"Serper ошибка: {e}"

    links, snippets = [], []
    for item in data.get("organic", []):
        link = item.get("link")
        if link and not is_bad_url(link):
            links.append(link)
            snippets.append((item.get("title", "") + " " + item.get("snippet", "")).strip())
    return links, snippets, "ok"


def search_duckduckgo(query: str, max_results: int = 8):
    urls, snippets = [], []
    for search_url in [
        "https://html.duckduckgo.com/html/?q=" + quote_plus(query),
        "https://duckduckgo.com/html/?q=" + quote_plus(query),
    ]:
        try:
            r = requests.get(search_url, headers=HEADERS, timeout=20)
            r.raise_for_status()
            soup = BeautifulSoup(r.text, "lxml")
            for a in soup.select("a.result__a, a.result-link, a[href]"):
                href = clean_ddg_url(a.get("href", ""))
                text = a.get_text(" ", strip=True)
                if href.startswith("http") and not is_bad_url(href):
                    urls.append(href)
                    snippets.append(text)
                if len(urls) >= max_results:
                    return urls, snippets, "ok"
        except Exception:
            continue
    return urls, snippets, "DuckDuckGo не дал ссылок"


def find_sources(product_name: str, category: str, serper_key: str, max_pages: int = 12):
    all_links, all_snippets, logs = [], [], []
    for q in build_queries(product_name, category):
        s_links, s_snips, s_status = search_serper(q, serper_key, 10)
        all_links.extend(s_links)
        all_snippets.extend(s_snips)
        logs.append(f"Serper: {q} | {len(s_links)} | {s_status}")

        d_links, d_snips, d_status = search_duckduckgo(q, 8)
        all_links.extend(d_links)
        all_snippets.extend(d_snips)
        logs.append(f"DuckDuckGo: {q} | {len(d_links)} | {d_status}")

    unique, seen = [], set()
    for url in all_links:
        if url not in seen and not is_bad_url(url):
            seen.add(url)
            unique.append(url)

    ranked = sorted(unique, key=lambda u: score_url(u, product_name), reverse=True)
    return ranked[:max_pages], all_snippets[:60], logs


def fetch_text(url: str):
    if is_bad_url(url):
        return "", "blacklist/bad domain"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
    except Exception as e:
        return "", f"не открылось: {e}"

    html = r.text or ""
    low = html.lower()
    if any(x in low for x in ["captcha", "recaptcha", "cloudflare", "access denied", "докажите", "робот", "checking your browser"]):
        return "", "капча/антибот, пропущено"

    try:
        soup = BeautifulSoup(html, "lxml")
        for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
            tag.decompose()
        text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
        return text[:12000], "ok"
    except Exception as e:
        return "", f"ошибка чтения страницы: {e}"
