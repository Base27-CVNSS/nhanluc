#!/usr/bin/env python3
"""Synchronize the 18 public documents for Tờ trình 617/TTr-UBND.

Official source:
https://hopkhonggiay.vinhlong.dcs.vn/Pages/Qr.aspx?Id=16834

The public page loads item (24) dynamically from document group 18725. This script
calls that same public endpoint, downloads the 18 attachments in their official order,
normalizes filenames to 01-18, extracts searchable text, and records SHA-256 hashes.
"""
from __future__ import annotations

import hashlib
import json
import sys
import urllib.parse
import urllib.request
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path

from docx import Document
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
FILES_DIR = ROOT / "files"
TEXT_DIR = ROOT / "text"
REPORT_PATH = ROOT / "sync_report.json"

BASE = "https://hopkhonggiay.vinhlong.dcs.vn"
MEETING_URL = f"{BASE}/Pages/Qr.aspx?Id=16834"
GROUP_URL = (
    f"{BASE}/Pages/Frontend/Mobile/LoadDanhSachChiaSeTaiLieu.aspx"
    "?pageId=1&pageSize=50&IDLichHop=16834"
)
GROUP_ID = "18725"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/150 Safari/537.36"


@dataclass(frozen=True)
class Spec:
    index: int
    output_name: str


SPECS = [
    Spec(1, "01_To_trinh_UBND_tinh_trinh_du_thao_Nghi_quyet.pdf"),
    Spec(2, "02_Du_thao_Nghi_quyet_trinh_thong_qua_HDND_tinh.docx"),
    Spec(3, "03_BC_tong_ket_cong_tac_dao_tao_chinh_sach_2021_2025.pdf"),
    Spec(4, "04_Bang_so_sanh_thuyet_minh_chinh_sach.pdf"),
    Spec(5, "05_BC_tong_hop_y_kien_gop_y_lan_1.pdf"),
    Spec(6, "06_BC_tong_hop_y_kien_gop_y_lan_2.pdf"),
    Spec(7, "07_4_2_BC_tong_hop_y_kien_gop_y_lan_2.pdf"),
    Spec(8, "08_05_BB_Hop_thong_nhat_noi_dung.pdf"),
    Spec(9, "09_BC_danh_gia_tac_dong_chinh_sach_dao_tao_nhan_luc.pdf"),
    Spec(10, "10_BC_tong_hop_gop_y_lan_3_chinh_sua.pdf"),
    Spec(11, "11_Bao_cao_Tham_dinh_Nghi_quyet.pdf"),
    Spec(12, "12_BC_tiep_thu_giai_trinh_tham_dinh_So_Tu_phap.pdf"),
    Spec(13, "13_Bao_cao_tham_tra_Nghi_quyet_HĐND.pdf"),
    Spec(14, "14_BC_tong_hop_y_kien_Thanh_vien_UBND_tinh.pdf"),
    Spec(15, "15_BC_giai_trinh_y_kien_Thanh_vien_UBND_tinh.pdf"),
    Spec(16, "16_TTr_trinh_ky_thong_qua_lai_du_thao_NQ.pdf"),
    Spec(17, "17_Thong_bao_ket_qua_dang_tai_du_thao_Nghi_quyet.pdf"),
    Spec(18, "18_811_BC_UBND_tiep_thu_giai_trinh_tham_tra.pdf"),
]


class LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[dict[str, str]] = []
        self._current: dict[str, str] | None = None
        self._buf: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() == "a":
            data = dict(attrs)
            self._current = {"href": data.get("href", ""), "text": ""}
            self._buf = []

    def handle_data(self, data: str) -> None:
        if self._current is not None:
            self._buf.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "a" and self._current is not None:
            self._current["text"] = " ".join("".join(self._buf).split())
            self.links.append(self._current)
            self._current = None
            self._buf = []


def request(url: str, *, data: bytes | None = None, referer: str = MEETING_URL) -> bytes:
    headers = {
        "User-Agent": USER_AGENT,
        "Referer": referer,
        "Accept": "*/*",
    }
    if data is not None:
        headers.update(
            {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With": "XMLHttpRequest",
            }
        )
    req = urllib.request.Request(url, data=data, headers=headers)
    with urllib.request.urlopen(req, timeout=90) as response:
        return response.read()


def get_official_links() -> list[dict[str, str]]:
    payload = urllib.parse.urlencode(
        {
            "tenFile": "",
            "nhomFileID": GROUP_ID,
            "isVaiTroChuanBi": "false",
            "isVaiTroKiemDuyet": "false",
        }
    ).encode("utf-8")
    html = request(GROUP_URL, data=payload).decode("utf-8", "replace")
    parser = LinkParser()
    parser.feed(html)
    links = [
        item
        for item in parser.links
        if "Command=DownloadFileByURL" in item.get("href", "")
    ]
    if len(links) != 18:
        raise RuntimeError(f"Expected 18 official attachment links, found {len(links)}")
    return links


def absolute_download_url(href: str) -> str:
    # The official href contains spaces inside the `url=` query value. Quote spaces and
    # non-ASCII characters while preserving the URL's structural delimiters.
    safe_href = urllib.parse.quote(href, safe="/:?=&%()+,.-_~")
    return urllib.parse.urljoin(BASE, safe_href)


def validate_binary(path: Path) -> None:
    data = path.read_bytes()
    if len(data) < 1000:
        raise ValueError(f"Downloaded file is unexpectedly small: {path.name} ({len(data)} bytes)")
    ext = path.suffix.lower()
    if ext == ".pdf" and not data.startswith(b"%PDF"):
        raise ValueError(f"Expected PDF signature for {path.name}")
    if ext == ".docx" and not data.startswith(b"PK"):
        raise ValueError(f"Expected DOCX/ZIP signature for {path.name}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def extract_text(path: Path) -> str:
    try:
        if path.suffix.lower() == ".pdf":
            reader = PdfReader(str(path))
            pages: list[str] = []
            for page_no, page in enumerate(reader.pages, 1):
                text = page.extract_text() or ""
                pages.append(f"\n--- TRANG {page_no} ---\n{text.strip()}\n")
            result = "".join(pages).strip()
            return result or "[PDF dạng ảnh hoặc không có lớp văn bản để trích xuất tự động.]"
        if path.suffix.lower() == ".docx":
            doc = Document(str(path))
            lines = [paragraph.text for paragraph in doc.paragraphs]
            for table in doc.tables:
                for row in table.rows:
                    lines.append("\t".join(cell.text for cell in row.cells))
            return "\n".join(lines).strip()
    except Exception as exc:
        return f"[Không thể trích xuất văn bản tự động: {exc}]"
    return ""


def main() -> int:
    FILES_DIR.mkdir(parents=True, exist_ok=True)
    TEXT_DIR.mkdir(parents=True, exist_ok=True)

    # Avoid stale files making an incomplete sync look complete.
    for old in FILES_DIR.iterdir():
        if old.is_file():
            old.unlink()
    for old in TEXT_DIR.glob("*.txt"):
        old.unlink()

    failures: list[dict] = []
    results: list[dict] = []

    try:
        links = get_official_links()
    except Exception as exc:
        REPORT_PATH.write_text(
            json.dumps(
                {
                    "meeting_url": MEETING_URL,
                    "group_id": GROUP_ID,
                    "expected": 18,
                    "downloaded": 0,
                    "failures": [{"stage": "list", "error": str(exc)}],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        raise

    for spec, link in zip(SPECS, links, strict=True):
        dest = FILES_DIR / spec.output_name
        source_name = link.get("text", "").strip()
        source_href = link.get("href", "")
        download_url = absolute_download_url(source_href)

        try:
            body = request(download_url)
            dest.write_bytes(body)
            validate_binary(dest)

            digest = sha256(dest)
            text_path = TEXT_DIR / f"{spec.index:02d}.txt"
            header = (
                f"TÀI LIỆU {spec.index:02d}/18\n"
                f"Tên hiển thị tại nguồn: {source_name}\n"
                f"Tệp chuẩn hóa: {spec.output_name}\n"
                f"Nguồn kỳ họp: {MEETING_URL}\n"
                f"Nhóm tài liệu: {GROUP_ID}\n"
                f"URL tải nguồn: {download_url}\n"
                f"SHA-256: {digest}\n"
                f"{'=' * 78}\n\n"
            )
            text_path.write_text(header + extract_text(dest) + "\n", encoding="utf-8")

            results.append(
                {
                    "index": spec.index,
                    "source_name": source_name,
                    "normalized_file": f"files/{spec.output_name}",
                    "searchable_text": f"text/{spec.index:02d}.txt",
                    "source_url": download_url,
                    "bytes": dest.stat().st_size,
                    "sha256": digest,
                }
            )
            print(f"[{spec.index:02d}/18] OK {source_name} -> {spec.output_name}")
        except Exception as exc:
            if dest.exists():
                dest.unlink()
            failures.append(
                {
                    "index": spec.index,
                    "source_name": source_name,
                    "source_url": download_url,
                    "error": str(exc),
                }
            )
            print(f"[{spec.index:02d}/18] ERROR {source_name}: {exc}", file=sys.stderr)

    report = {
        "meeting_url": MEETING_URL,
        "attachment_endpoint": GROUP_URL,
        "group_id": GROUP_ID,
        "expected": 18,
        "downloaded": len(results),
        "results": results,
        "failures": failures,
    }
    REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    if failures or len(results) != 18:
        print(f"ERROR: synchronized {len(results)}/18 documents", file=sys.stderr)
        return 2
    print("OK: synchronized and verified all 18 official documents")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
