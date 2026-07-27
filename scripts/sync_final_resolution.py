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
OUT_DIR = Path(__file__).resolve().parents[1] / "van-ban-chinh-thuc"
OUT_FILE = OUT_DIR / "47_2026_NQ_HDND.pdf"
META_FILE = OUT_DIR / "47_2026_NQ_HDND.metadata.json"

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


def candidate_links(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    links: list[str] = []
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        label = " ".join(a.stripped_strings)
        hay = unquote(f"{href} {label}").lower()
        if ".pdf" in hay and ("47_2026" in hay or "47/2026" in hay or "nq hdnd" in hay):
            links.append(urljoin(PAGE_URL, href))

    # Some CMS templates render attachment URLs inside scripts/data attributes.
    for raw in re.findall(r"(?:https?://[^\"'<>\s]+|/[^\"'<>\s]+\.pdf[^\"'<>\s]*)", html, flags=re.I):
        hay = unquote(raw).lower()
        if "47_2026" in hay or "47%5f2026" in hay or "47%202026" in hay:
            links.append(urljoin(PAGE_URL, raw.replace("&amp;", "&")))

    seen = set()
    ordered = []
    for link in links:
        if link not in seen:
            seen.add(link)
            ordered.append(link)
    return ordered


def main() -> None:
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/150 Safari/537.36",
        "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
    })

    page = session.get(PAGE_URL, timeout=60)
    page.raise_for_status()
    links = candidate_links(page.text)
    if not links:
        raise RuntimeError("Không tìm thấy liên kết PDF đính kèm của Nghị quyết 47/2026/NQ-HĐND trên trang nguồn.")

    pdf_url = None
    pdf_data = None
    errors = []
    for link in links:
        try:
            resp = session.get(link, timeout=90, allow_redirects=True)
            resp.raise_for_status()
            data = resp.content
            if data[:5] == b"%PDF-" and len(data) > 10_000:
                pdf_url = resp.url
                pdf_data = data
                break
            errors.append(f"{link}: not a PDF ({resp.headers.get('content-type')}, {len(data)} bytes)")
        except Exception as exc:
            errors.append(f"{link}: {exc}")

    if pdf_data is None or pdf_url is None:
        raise RuntimeError("Không tải được PDF chính thức. " + " | ".join(errors[:5]))

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_bytes(pdf_data)
    meta = {
        **EXPECTED,
        "source_file_url": pdf_url,
        "repository_file": str(OUT_FILE.relative_to(OUT_FILE.parents[1])).replace("\\", "/"),
        "bytes": len(pdf_data),
        "sha256": sha256(pdf_data),
    }
    META_FILE.write_text(json.dumps(meta, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
