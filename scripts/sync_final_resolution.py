#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from urllib.parse import urljoin, unquote

import requests
from bs4 import BeautifulSoup

PAGE_URL = "https://hdnd.vinhlong.gov.vn/chi-tiet-van-ban/VB_ID=16778"
# Public mirror fallback discovered from a legal-document index. It is accepted only
# when its SHA-256 exactly matches the 12-page official PDF supplied for verification.
MIRROR_PDF_URL = "https://lsu.vn/api/tai-ve?doc=193309&f=47_2026_NQ_H%C4%90ND_0001.signed.pdf"
EXPECTED_PDF_SHA256 = "9e122c0ac0161d7fef2341a3420a93f740eddc5e6ab088d6bb69b0ee2b9848d3"

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "van-ban-chinh-thuc"
OUT_FILE = OUT_DIR / "47_2026_NQ_HDND.pdf"
META_FILE = OUT_DIR / "47_2026_NQ_HDND.metadata.json"
DIAG_FILE = OUT_DIR / "source_diagnostic.json"

EXPECTED = {
    "so_ky_hieu": "47/2026/NQ-HĐND",
    "ngay_ban_hanh": "15/07/2026",
    "ngay_co_hieu_luc": "25/07/2026",
    "trich_yeu": "Nghị quyết Ban hành Quy định chính sách hỗ trợ đào tạo cán bộ, công chức; thu hút nguồn nhân lực chất lượng cao tỉnh Vĩnh Long giai đoạn 2026 - 2030",
    "loai_van_ban": "Nghị Quyết",
    "co_quan_ban_hanh": "HĐND tỉnh Vĩnh Long",
    "nguoi_ky": "Nguyễn Minh Dũng",
    "chuc_vu": "Chủ tịch",
    "official_source_page": PAGE_URL,
}


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def collect_links(html: str) -> list[dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    items: list[dict[str, str]] = []
    for tag in soup.find_all(["a", "iframe", "embed", "object", "source"]):
        raw = tag.get("href") or tag.get("src") or tag.get("data") or ""
        if not raw:
            continue
        items.append({
            "tag": tag.name,
            "label": " ".join(tag.stripped_strings)[:300],
            "raw": raw,
            "absolute": urljoin(PAGE_URL, raw),
        })
    return items


def candidate_links(html: str, all_links: list[dict[str, str]]) -> list[str]:
    candidates: list[str] = []
    for item in all_links:
        hay = unquote(f"{item['raw']} {item['label']}").lower()
        if any(token in hay for token in [".pdf", "47_2026", "47 2026", "47/2026", "nq hdnd", "download", "file"]):
            candidates.append(item["absolute"])

    patterns = [
        r"https?://[^\"'<>\s]+",
        r"/[^\"'<>\s]+\.pdf(?:\?[^\"'<>\s]*)?",
        r"/[^\"'<>\s]*(?:Download|download|TaiFile|File|file)[^\"'<>\s]*",
    ]
    for pattern in patterns:
        for raw in re.findall(pattern, html, flags=re.I):
            hay = unquote(raw).lower()
            if any(token in hay for token in ["47_2026", "47%5f2026", "47%202026", "47/2026", ".pdf", "download", "file"]):
                candidates.append(urljoin(PAGE_URL, raw.replace("&amp;", "&")))

    seen: set[str] = set()
    ordered: list[str] = []
    for link in candidates:
        if link not in seen:
            seen.add(link)
            ordered.append(link)
    return ordered


def try_pdf(session: requests.Session, url: str) -> tuple[bytes | None, str, str]:
    try:
        resp = session.get(url, timeout=90, allow_redirects=True)
        resp.raise_for_status()
        data = resp.content
        ctype = (resp.headers.get("content-type") or "").lower()
        if data[:5] == b"%PDF-" and len(data) > 10_000:
            return data, resp.url, ""
        return None, resp.url, f"not PDF ({ctype}, {len(data)} bytes)"
    except Exception as exc:
        return None, url, str(exc)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/150 Safari/537.36",
        "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })

    links: list[str] = []
    diagnostic: dict = {"official_page": PAGE_URL, "expected_sha256": EXPECTED_PDF_SHA256}
    try:
        page = session.get(PAGE_URL, timeout=60, allow_redirects=True)
        page.raise_for_status()
        all_links = collect_links(page.text)
        links = candidate_links(page.text, all_links)
        diagnostic.update({
            "official_page_final_url": page.url,
            "official_page_status": page.status_code,
            "official_page_content_type": page.headers.get("content-type"),
            "official_page_html_bytes": len(page.content),
            "official_candidate_links": links,
        })
    except Exception as exc:
        diagnostic["official_page_error"] = str(exc)

    attempts = list(links)
    if MIRROR_PDF_URL not in attempts:
        attempts.append(MIRROR_PDF_URL)

    pdf_data = None
    pdf_url = None
    method = None
    errors: list[str] = []
    for link in attempts:
        data, final_url, err = try_pdf(session, link)
        if data is None:
            errors.append(f"{link}: {err}")
            continue

        digest = sha256(data)
        # The mirror is a technical fallback only: never accept it unless byte-for-byte
        # identical to the verified 12-page official PDF.
        if link == MIRROR_PDF_URL and digest != EXPECTED_PDF_SHA256:
            errors.append(f"mirror SHA mismatch: got {digest}")
            continue

        pdf_data = data
        pdf_url = final_url
        method = "official_page_attachment" if link != MIRROR_PDF_URL else "verified_public_mirror"
        break

    diagnostic["attempts"] = attempts
    diagnostic["errors"] = errors
    DIAG_FILE.write_text(json.dumps(diagnostic, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if pdf_data is None or pdf_url is None or method is None:
        raise RuntimeError("Không tải được PDF đã xác minh; xem van-ban-chinh-thuc/source_diagnostic.json")

    OUT_FILE.write_bytes(pdf_data)
    digest = sha256(pdf_data)
    meta = {
        **EXPECTED,
        "download_method": method,
        "download_url": pdf_url,
        "repository_file": "van-ban-chinh-thuc/47_2026_NQ_HDND.pdf",
        "bytes": len(pdf_data),
        "sha256": digest,
        "verified_against_uploaded_official_pdf": digest == EXPECTED_PDF_SHA256,
    }
    META_FILE.write_text(json.dumps(meta, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
