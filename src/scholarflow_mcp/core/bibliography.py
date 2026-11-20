"""Stage 6: Live Crossref/OpenAlex DOI verification & clean Mendeley/Zotero BibTeX/RIS export."""

from __future__ import annotations

import re
import urllib.parse
from pathlib import Path
from typing import Any

import httpx
import openpyxl

from scholarflow_mcp.config import ProjectConfig
from scholarflow_mcp.core.records import normalize_doi

USER_AGENT = "ScholarFlowMCP/0.1.0 (https://github.com/AkmalRamadhan04/scholarflow-mcp; academic-slr)"


def split_authors_clean(raw_authors: str) -> list[str]:
    """Split multi-author strings cleanly into individual authors (eliminates RIS `&` author bugs)."""
    if not raw_authors:
        return ["Anonymous"]
    s = str(raw_authors).strip()
    # Priority 1: semicolon-separated (standard Scopus/WoS/RIS)
    if ";" in s:
        parts = [p.strip().lstrip("&").strip() for p in s.split(";") if p.strip()]
        return [p for p in parts if p] or ["Anonymous"]
    # Priority 2: ' and ' or ' & ' separated
    s = re.sub(r"\s+&\s+", " and ", s)
    if " and " in s:
        parts = [p.strip().rstrip(",").strip() for p in re.split(r"\s+and\s+", s) if p.strip()]
        return [p for p in parts if p] or ["Anonymous"]
    return [s]


def make_citekey(study_code: str, authors: list[str], year: str) -> str:
    first = authors[0] if authors else "anon"
    surname = first.split(",")[0].strip().split()[-1] if first else "anon"
    surname_clean = re.sub(r"[^a-zA-Z0-9]", "", surname).lower() or "anon"
    yr = re.sub(r"[^0-9]", "", str(year or ""))[:4] or "nd"
    code_clean = re.sub(r"[^a-zA-Z0-9]", "", study_code).lower()
    return f"{surname_clean}{yr}_{code_clean}" if code_clean else f"{surname_clean}{yr}"


def _verify_doi_online(client: httpx.Client, doi: str) -> dict[str, Any]:
    try:
        r = client.get(f"https://api.crossref.org/works/{urllib.parse.quote(doi)}")
        if r.status_code == 200:
            msg = r.json().get("message", {})
            titles = msg.get("title") or []
            return {
                "doi": doi,
                "verified": True,
                "provider": "crossref",
                "remote_title": titles[0] if titles else "",
            }
    except Exception:
        pass
    try:
        r2 = client.get(f"https://api.openalex.org/works/https://doi.org/{urllib.parse.quote(doi)}")
        if r2.status_code == 200:
            obj = r2.json()
            return {
                "doi": doi,
                "verified": True,
                "provider": "openalex",
                "remote_title": obj.get("title", ""),
            }
    except Exception:
        pass
    return {"doi": doi, "verified": False, "provider": "none", "remote_title": ""}


def verify_and_export_bibliography(
    project_dir: str,
    verify_online: bool = False,
    extra_references: list[dict[str, str]] | None = None,
    timeout_seconds: float = 10.0,
) -> dict[str, Any]:
    """Export clean, validated BibTeX (.bib) and RIS (.ris) files from Extraction_Matrix.xlsx (+ extra refs)."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()
    matrix_path = cfg.resolve_path(root, "extraction_matrix_file")
    bib_path = cfg.resolve_path(root, "bib_file")
    ris_path = cfg.resolve_path(root, "ris_file")

    wb = openpyxl.load_workbook(matrix_path, data_only=True)
    ws = wb[preset.sheet_matrix]
    headers = [c.value for c in ws[1]]
    col = {h: i for i, h in enumerate(headers) if h}

    entries: list[dict[str, Any]] = []
    for row in list(ws.iter_rows(values_only=True))[1:]:
        if not row or not row[0]:
            continue
        code = str(row[col["Study_Code"]]).strip()
        authors_list = split_authors_clean(str(row[col.get("Authors", 1)] or ""))
        year = str(row[col.get("Year", 2)] or "n.d.").strip()[:4]
        title = str(row[col.get("Title", 3)] or "").strip()
        venue = str(row[col.get("Venue", 5)] or "").strip()
        doi = normalize_doi(str(row[col.get("DOI", 6)] or ""))
        entries.append(
            {
                "study_code": code,
                "citekey": make_citekey(code, authors_list, year),
                "authors": authors_list,
                "year": year,
                "title": title,
                "venue": venue,
                "doi": doi,
            }
        )

    for idx, extra in enumerate(extra_references or [], start=1):
        authors_list = split_authors_clean(str(extra.get("authors", "")))
        year = str(extra.get("year", "n.d.")).strip()[:4]
        code = str(extra.get("study_code", f"REF{idx:02d}"))
        entries.append(
            {
                "study_code": code,
                "citekey": extra.get("citekey") or make_citekey(code, authors_list, year),
                "authors": authors_list,
                "year": year,
                "title": str(extra.get("title", "")).strip(),
                "venue": str(extra.get("venue", "")).strip(),
                "doi": normalize_doi(extra.get("doi", "")),
            }
        )

    doi_results: list[dict[str, Any]] = []
    if verify_online:
        with httpx.Client(headers={"User-Agent": USER_AGENT}, timeout=timeout_seconds) as client:
            for e in entries:
                if e["doi"]:
                    doi_results.append(_verify_doi_online(client, e["doi"]))

    bib_blocks: list[str] = []
    ris_lines: list[str] = []

    for e in entries:
        bib_auth = " and ".join(e["authors"])
        b_lines = [
            f"@article{{{e['citekey']},",
            f"  author  = {{{bib_auth}}},",
            f"  year    = {{{e['year']}}},",
            f"  title   = {{{e['title']}}},",
        ]
        if e["venue"]:
            b_lines.append(f"  journal = {{{e['venue']}}},")
        if e["doi"]:
            b_lines.append(f"  doi     = {{{e['doi']}}},")
        b_lines.append(f"  note    = {{Corpus Code: {e['study_code']}}}")
        b_lines.append("}")
        bib_blocks.append("\n".join(b_lines))

        ris_lines.append("TY  - JOUR")
        for a in e["authors"]:
            # Guarantee zero '&' author joins inside RIS AU lines
            clean_a = a.replace("&", "").strip()
            ris_lines.append(f"AU  - {clean_a}")
        ris_lines.append(f"PY  - {e['year']}")
        ris_lines.append(f"TI  - {e['title']}")
        if e["venue"]:
            ris_lines.append(f"JO  - {e['venue']}")
        if e["doi"]:
            ris_lines.append(f"DO  - {e['doi']}")
        ris_lines.append(f"N1  - Corpus Code: {e['study_code']}")
        ris_lines.append("ER  - ")
        ris_lines.append("")

    bib_path.parent.mkdir(parents=True, exist_ok=True)
    bib_path.write_text("\n\n".join(bib_blocks) + ("\n" if bib_blocks else ""), encoding="utf-8")
    ris_path.write_text("\n".join(ris_lines), encoding="utf-8")

    return {
        "status": "ok",
        "entries_exported": len(entries),
        "bib_path": str(bib_path),
        "ris_path": str(ris_path),
        "online_verification_ran": verify_online,
        "verified_dois_count": sum(1 for d in doi_results if d["verified"]),
        "unverified_dois": [d["doi"] for d in doi_results if not d["verified"]],
    }
