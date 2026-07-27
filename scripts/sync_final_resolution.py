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
    "source_page": PAGE_URL,
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

    # Also inspect quoted URLs and server-relative file paths embedded in scripts/JSON.
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


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/150 Safari/537.36",
        "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": "https://hdnd.vinhlong.gov.vn/",
    })

    page = session.get(PAGE_URL, timeout=60, allow_redirects=True)
    page.raise_for_status()
    all_links = collect_links(page.text)
    links = candidate_links(page.text, all_links)

    diagnostic = {
        "requested_url": PAGE_URL,
        "final_url": page.url,
        "status": page.status_code,
        "content_type": page.headers.get("content-type"),
        "html_bytes": len(page.content),
        "page_title": BeautifulSoup(page.text, "html.parser").title.string.strip() if BeautifulSoup(page.text, "html.parser").title and BeautifulSoup(page.text, "html.parser").title.string else "",
        "candidate_links": links,
        "all_links": all_links,
        "html_preview": page.text[:12000],
    }
    DIAG_FILE.write_text(json.dumps(diagnostic, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if not links:
        raise RuntimeError("Không tìm thấy liên kết tệp ứng viên; xem van-ban-chinh-thuc/source_diagnostic.json")

    pdf_url = None
    pdf_data = None
    errors: list[str] = []
    for link in links:
        try:
            resp = session.get(link, timeout=90, allow_redirects=True)
            resp.raise_for_status()
            data = resp.content
            ctype = (resp.headers.get("content-type") or "").lower()
            if data[:5] == b"%PDF-" and len(data) > 10_000:
                pdf_url = resp.url
                pdf_data = data
                break
            errors.append(f"{link}: not PDF ({ctype}, {len(data)} bytes, final={resp.url})")
        except Exception as exc:
            errors.append(f"{link}: {exc}")

    diagnostic["download_errors"] = errors
    DIAG_FILE.write_text(json.dumps(diagnostic, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if pdf_data is None or pdf_url is None:
        raise RuntimeError("Không tải được PDF chính thức; xem source_diagnostic.json")

    OUT_FILE.write_bytes(pdf_data)
    meta = {
        **EXPECTED,
        "source_file_url": pdf_url,
        "repository_file": "van-ban-chinh-thuc/47_2026_NQ_HDND.pdf",
        "bytes": len(pdf_data),
        "sha256": sha256(pdf_data),
    }
    META_FILE.write_text(json.dumps(meta, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
