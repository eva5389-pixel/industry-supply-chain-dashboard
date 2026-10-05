import ipaddress
import re
import socket
import ssl
from html import unescape
from urllib.parse import parse_qs, unquote, urlencode, urljoin, urlparse, urlunparse

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3 import PoolManager


LABEL_PATTERN = re.compile(
    r"(?:Benchmark|比較基準|參考指標|指標指數|基準指數|追蹤指數)\s*[:：\-]?\s*"
    r"([^\n\r|｜<>]{2,80}?(?:指數|Index))",
    flags=re.IGNORECASE,
)
INDEX_PATTERN = re.compile(
    r"([A-Za-z0-9&／/・.\-\s\u4e00-\u9fff]{2,60}(?:指數|Index))",
    flags=re.IGNORECASE,
)
FUND_PATTERN = re.compile(
    r"([A-Za-z0-9Ａ-Ｚａ-ｚ０-９&（）()／/・.\-\s\u4e00-\u9fff]{2,80}基金)"
)


class _MoneyDJTLSAdapter(HTTPAdapter):
    """Verify TLS while tolerating MoneyDJ's legacy certificate chain."""

    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        context = ssl.create_default_context()
        if hasattr(ssl, "VERIFY_X509_STRICT"):
            context.verify_flags &= ~ssl.VERIFY_X509_STRICT
        pool_kwargs["ssl_context"] = context
        self.poolmanager = PoolManager(
            num_pools=connections,
            maxsize=maxsize,
            block=block,
            **pool_kwargs,
        )


def _resolve_moneydj_wrapper_url(url):
    parsed = urlparse(url)
    if not parsed.hostname or not parsed.hostname.lower().endswith(".moneydj.com"):
        return url
    if parsed.path.rstrip("/").lower() != "/main.html":
        return url
    route = unquote(parse_qs(parsed.query).get("sUrl", [""])[0])
    page_match = re.search(r"\$WR(\d{2})\]DJHTM", route, flags=re.IGNORECASE)
    fund_match = re.search(r"\{A\}([A-Z0-9]+(?:-[A-Z0-9]+)?)", route, flags=re.IGNORECASE)
    if not page_match or not fund_match:
        return url
    return urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            f"/w/wr/wr{page_match.group(1)}.djhtm",
            "",
            urlencode({"a": fund_match.group(1).upper()}),
            "",
        )
    )


def _validate_public_url(url):
    parsed = urlparse(url.strip())
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError("請貼上完整的 http 或 https 基金網址。")
    if parsed.username or parsed.password:
        raise ValueError("網址不可包含帳號或密碼。")
    try:
        addresses = socket.getaddrinfo(
            parsed.hostname, parsed.port or (443 if parsed.scheme == "https" else 80)
        )
    except socket.gaierror as exc:
        raise ValueError("無法解析網址主機。") from exc
    if any(not ipaddress.ip_address(item[4][0]).is_global for item in addresses):
        raise ValueError("基於安全考量，不接受內網或本機網址。")
    return parsed.geturl()


def _get_page(url, **kwargs):
    hostname = (urlparse(url).hostname or "").lower()
    if hostname.endswith(".moneydj.com"):
        session = requests.Session()
        session.mount("https://", _MoneyDJTLSAdapter())
        return session.get(url, **kwargs)
    return requests.get(url, **kwargs)


def _download_html(url, max_bytes=8_000_000):
    safe_url = _validate_public_url(url)
    for _ in range(4):
        response = _get_page(
            safe_url,
            timeout=(15, 35),
            headers={"User-Agent": "Mozilla/5.0 IndustrySupplyChainDashboard/1.0"},
            allow_redirects=False,
            stream=True,
        )
        if response.is_redirect or response.is_permanent_redirect:
            destination = response.headers.get("location")
            response.close()
            if not destination:
                raise ValueError("網站重新導向缺少目的網址。")
            safe_url = _validate_public_url(urljoin(safe_url, destination))
            continue
        response.raise_for_status()
        chunks = []
        size = 0
        for chunk in response.iter_content(64 * 1024):
            size += len(chunk)
            if size > max_bytes:
                raise ValueError("基金網頁超過 8 MB，請改用手動填寫。")
            chunks.append(chunk)
        encoding = response.encoding or response.apparent_encoding or "utf-8"
        return b"".join(chunks).decode(encoding, errors="replace"), response.url
    raise ValueError("網站重新導向次數過多。")


def _clean(value):
    return re.sub(r"\s+", " ", unescape(value)).strip(" -–—|｜：:")


def extract_fund_metadata(html, source_url=""):
    soup = BeautifulSoup(html, "html.parser")
    text = _clean(soup.get_text("\n", strip=True))
    script_text = "\n".join(
        script.get_text(" ", strip=True) for script in soup.find_all("script")
    )
    title_candidates = []
    for selector in ("h1", "h2", "h3", "h4"):
        title_candidates.extend(
            _clean(node.get_text(" ", strip=True)) for node in soup.select(selector)
        )
    for key, value in (("property", "og:title"), ("name", "twitter:title")):
        node = soup.find("meta", attrs={key: value})
        if node and node.get("content"):
            title_candidates.append(_clean(node["content"]))
    if soup.title and soup.title.string:
        title_candidates.append(_clean(soup.title.string))
    title_candidates.append(text[:500])

    names = []
    for candidate in title_candidates:
        for match in FUND_PATTERN.finditer(candidate):
            name = re.sub(r"基金(?:\s+基金)+$", "基金", _clean(match.group(1)))
            if name not in names:
                names.append(name)
    generic = {"基金", "國內基金", "境外基金", "海外基金", "單一基金", "基金資訊"}
    names = [name for name in names if name not in generic]
    fund_name = max(
        names,
        key=lambda name: (bool(re.match(r"^\d", name)), len(name)),
        default="",
    )

    combined = "\n".join((text, script_text, html))
    labelled = LABEL_PATTERN.search(combined)
    benchmark = _clean(labelled.group(1)) if labelled else ""
    if not benchmark:
        candidates = []
        for match in INDEX_PATTERN.finditer(combined):
            value = _clean(match.group(1))
            if any(noise in value for noise in ("基金績效", "指數型基金", "指數基金")):
                continue
            if value not in candidates:
                candidates.append(value)
        if len(candidates) == 1:
            benchmark = candidates[0]
    return {"fund_name": fund_name, "benchmark": benchmark, "source_url": source_url}


def fetch_fund_metadata(url):
    resolved_url = _resolve_moneydj_wrapper_url(unquote(url))
    html, final_url = _download_html(resolved_url)
    result = extract_fund_metadata(html, final_url)
    if not result["fund_name"] and not result["benchmark"]:
        raise ValueError("這個頁面沒有可辨識的基金名稱或 Benchmark，請手動填寫。")
    return result
