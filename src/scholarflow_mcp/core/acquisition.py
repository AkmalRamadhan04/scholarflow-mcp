"""Stage 2: Legal Open-Access PDF retriever (strict CAPTCHA refusal) & clickable manual download queue."""

from __future__ import annotations

import re
import urllib.parse
from pathlib import Path
from typing import Any

import httpx
import openpyxl
import pymupdf

from scholarflow_mcp.config import ProjectConfig
from scholarflow_mcp.core.excel_utils import LINK_FONT, style_sheet

CAPTCHA_MARKERS = (
    "captcha",
    "cf-turnstile",
    "verify you are human",
    "recaptcha",
    "hcaptcha",
    "cloudflare challenge",
    "attention required! | cloudflare",
)

USER_AGENT = "ScholarFlowMCP/0.1.0 (https://github.com/AkmalRamadhan04/scholarflow-mcp; academic-slr)"


def build_safe_filename(record_id: str, authors: str, year: str, title: str) -> str:
    """Create a standardized, OS-safe target PDF filename."""
    first_auth = str(authors or "Anonymous").split(";")[0].split(",")[0].strip()
    first_auth = re.sub(r'[\\/:*?"<>|]', "", first_auth)[:25].strip() or "Anonymous"
    yr = re.sub(r"[^0-9]", "", str(year or ""))[:4] or "nd"
    clean_title = re.sub(r'[\\/:*?"<>|]', "", str(title or ""))
    short_title = " ".join(clean_title.split()[:6]) or "Untitled"
    return f"{record_id}_{first_auth} ({yr}) - {short_title}.pdf"


def to_clickable_url(doi: str | None, url: str | None) -> str:
    for val in (url, doi):
        if not val:
            continue
        s = str(val).strip()
        if s in {"-", "—", ""}:
            continue
        if s.startswith(("http://", "https://")):
            return s
        if s.startswith("10."):
            return f"https://doi.org/{s}"
    return ""


def extract_pdf_url_from_html(html_text: str, base_url: str) -> tuple[str | None, bool]:
    """Extract an open-access PDF link from publisher HTML; returns (url, blocked_by_captcha)."""
    low = html_text.lower()
    if any(marker in low for marker in CAPTCHA_MARKERS):
        return None, True

    # 1. Google Scholar / OJS / DSpace / EPrints standard citation_pdf_url meta tag
    patterns = [
        r'<meta[^>]+name=["\']citation_pdf_url["\'][^>]+content=["\']([^"\']+)["\']',
        r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']citation_pdf_url["\']',
        r'href=["\']([^"\']*/article/download/[^"\']+)["\']',
    ]
    for pat in patterns:
        m = re.search(pat, html_text, re.IGNORECASE)
        if m:
            return urllib.parse.urljoin(base_url, m.group(1).replace("&amp;", "&")), False
    return None, False


def _is_valid_pdf_bytes(data: bytes) -> bool:
    if not data or not data[:8].startswith(b"%PDF"):
        return False
    try:
        doc = pymupdf.open(stream=data, filetype="pdf")
        ok = len(doc) >= 1
        doc.close()
        return ok
    except Exception:
        return False


def fetch_open_access_pdfs(
    project_dir: str,
    candidates: list[dict[str, str]],
    timeout_seconds: float = 15.0,
) -> dict[str, Any]:
    """Attempt legal OA PDF retrieval via OpenAlex API and HTML citation_pdf_url discovery.

    Strictly refuses/skips CAPTCHA or bot-gated pages and verifies every downloaded PDF with PyMuPDF.
    """
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    stage_dir = cfg.resolve_path(root, "unclassified_dir")
    stage_dir.mkdir(parents=True, exist_ok=True)

    downloaded: list[dict[str, str]] = []
    skipped_captcha: list[dict[str, str]] = []
    failed: list[dict[str, str]] = []

    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "application/pdf,text/html,application/xhtml+xml,*/*;q=0.8",
    }

    with httpx.Client(headers=headers, timeout=timeout_seconds, follow_redirects=True) as client:
        for item in candidates:
            rec_id = str(item.get("record_id", "REC"))
            doi = str(item.get("doi", "")).strip()
            url = str(item.get("url", "")).strip()
            target_name = build_safe_filename(
                rec_id,
                str(item.get("authors", "")),
                str(item.get("year", "")),
                str(item.get("title", "")),
            )
            target_path = stage_dir / target_name
            if target_path.exists():
                downloaded.append({"record_id": rec_id, "file": str(target_path), "note": "already_exists"})
                continue

            candidate_urls: list[str] = []
            if doi:
                try:
                    oa_api = f"https://api.openalex.org/works/https://doi.org/{urllib.parse.quote(doi)}"
                    r = client.get(oa_api)
                    if r.status_code == 200:
                        obj = r.json()
                        best = obj.get("best_oa_location") or {}
                        oa = obj.get("open_access") or {}
                        for u in (best.get("pdf_url"), oa.get("oa_url"), best.get("landing_page_url")):
                            if u and u not in candidate_urls:
                                candidate_urls.append(u)
                except Exception:
                    pass
            if url and url not in candidate_urls:
                candidate_urls.append(url)
            if doi:
                doi_url = f"https://doi.org/{doi}"
                if doi_url not in candidate_urls:
                    candidate_urls.append(doi_url)

            success = False
            hit_captcha = False
            last_reason = "no_open_access_pdf_found"

            for cand_url in candidate_urls:
                try:
                    resp = client.get(cand_url)
                    if resp.status_code >= 400:
                        last_reason = f"http_{resp.status_code}"
                        continue
                    ctype = resp.headers.get("content-type", "").lower()
                    if "pdf" in ctype or resp.content[:4] == b"%PDF":
                        if _is_valid_pdf_bytes(resp.content):
                            target_path.write_bytes(resp.content)
                            downloaded.append(
                                {"record_id": rec_id, "file": str(target_path), "source_url": str(resp.url)}
                            )
                            success = True
                            break
                        last_reason = "invalid_pdf_payload"
                    elif "html" in ctype:
                        pdf_url, is_captcha = extract_pdf_url_from_html(resp.text, str(resp.url))
                        if is_captcha:
                            hit_captcha = True
                            last_reason = "blocked_by_captcha_or_bot_gate"
                            break
                        if pdf_url:
                            r2 = client.get(pdf_url)
                            if r2.status_code == 200 and _is_valid_pdf_bytes(r2.content):
                                target_path.write_bytes(r2.content)
                                downloaded.append(
                                    {"record_id": rec_id, "file": str(target_path), "source_url": str(r2.url)}
                                )
                                success = True
                                break
                except Exception as exc:
                    last_reason = f"network_error: {type(exc).__name__}"

            if not success:
                entry = {"record_id": rec_id, "reason": last_reason}
                if hit_captcha:
                    skipped_captcha.append(entry)
                else:
                    failed.append(entry)

    return {
        "status": "ok",
        "staging_dir": str(stage_dir),
        "downloaded_count": len(downloaded),
        "skipped_captcha_count": len(skipped_captcha),
        "failed_count": len(failed),
        "downloaded": downloaded,
        "skipped_captcha": skipped_captcha,
        "failed": failed,
    }


def build_manual_download_queue(
    project_dir: str,
    queue_items: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    """Generate a prioritized manual download queue with clickable Excel hyperlinks in Search_Log.xlsx.

    If `queue_items` is omitted, automatically pulls all Candidates where
    TA_Decision in ('Include', 'Unclear', 'Y') and Fulltext_Status == 'Not_Retrieved'.
    """
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()
    log_path = cfg.resolve_path(root, "search_log_file")

    wb = openpyxl.load_workbook(log_path)
    items: list[dict[str, str]] = []

    if queue_items is not None:
        items = queue_items
    else:
        ws_cand = wb[preset.sheet_candidates]
        headers = [c.value for c in ws_cand[1]]
        col = {h: i for i, h in enumerate(headers) if h}
        for row in list(ws_cand.iter_rows(values_only=True))[1:]:
            if not row[0]:
                continue
            ta = str(row[col.get("TA_Decision", 9)] or "").strip().lower()
            ft = str(row[col.get("Fulltext_Status", 10)] or "").strip().lower()
            if ta in {"include", "unclear", "y", "yes", "pass"} and ft in {"not_retrieved", "t", "pending", ""}:
                items.append(
                    {
                        "record_id": str(row[col.get("Record_ID", 0)]),
                        "pathway": str(row[col.get("Pathway", 1)] or ""),
                        "priority": "High" if ta == "include" else "Medium",
                        "authors": str(row[col.get("Authors", 3)] or ""),
                        "year": str(row[col.get("Year", 4)] or ""),
                        "title": str(row[col.get("Title", 5)] or ""),
                        "venue": str(row[col.get("Venue", 6)] or ""),
                        "doi": str(row[col.get("DOI", 7)] or ""),
                        "url": str(row[col.get("URL", 8)] or ""),
                        "status": "Pending_Manual_Download",
                        "notes": "",
                    }
                )

    ws_q = wb[preset.sheet_download_queue]
    if ws_q.max_row > 1:
        ws_q.delete_rows(2, ws_q.max_row - 1)

    clickable_count = 0
    for idx, it in enumerate(items, start=1):
        qid = f"Q{idx:03d}"
        rec_id = str(it.get("record_id", qid))
        link = to_clickable_url(it.get("doi"), it.get("url"))
        suggested = build_safe_filename(
            rec_id,
            str(it.get("authors", "")),
            str(it.get("year", "")),
            str(it.get("title", "")),
        )
        ws_q.append(
            [
                qid,
                rec_id,
                str(it.get("pathway", "")),
                str(it.get("priority", "High")),
                str(it.get("authors", "")),
                str(it.get("year", "")),
                str(it.get("title", "")),
                str(it.get("venue", "")),
                str(it.get("doi", "")),
                suggested,
                link if link else "No URL/DOI available",
                str(it.get("status", "Pending_Manual_Download")),
                str(it.get("notes", "")),
            ]
        )
        row_num = ws_q.max_row
        if link:
            cell = ws_q.cell(row_num, 11)
            cell.hyperlink = link
            cell.font = LINK_FONT
            clickable_count += 1

    style_sheet(
        ws_q,
        {1: 10, 2: 12, 3: 22, 4: 10, 5: 24, 6: 8, 7: 42, 8: 26, 9: 26, 10: 40, 11: 42, 12: 24, 13: 28},
    )
    wb.save(log_path)

    return {
        "status": "ok",
        "queue_count": len(items),
        "clickable_links_count": clickable_count,
        "sheet": preset.sheet_download_queue,
        "search_log": str(log_path),
    }
