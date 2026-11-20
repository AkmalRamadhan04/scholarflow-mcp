"""Stage 1: Multi-database bibliographic ingestion, deduplication & SPAR-4-SLR / PRISMA reconciliation."""

from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any

import openpyxl

from scholarflow_mcp.config import ProjectConfig
from scholarflow_mcp.core.excel_utils import style_sheet

PATHWAY_KEYS = (
    "pathway_1_primary_db",
    "pathway_2_secondary_db",
    "pathway_3_web_complement",
    "pathway_4_team_seed",
    "pathway_5_snowballing",
)


def normalize_doi(raw: str | None) -> str:
    if not raw:
        return ""
    s = str(raw).strip().lower()
    s = re.sub(r"^https?://(dx\.)?doi\.org/", "", s)
    s = re.sub(r"^doi:\s*", "", s)
    return s.strip().rstrip(".")


def normalize_title(raw: str | None) -> str:
    if not raw:
        return ""
    s = str(raw).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _pick(row: dict[str, Any], candidates: tuple[str, ...]) -> str:
    lower_map = {k.strip().lower().lstrip("\ufeff"): v for k, v in row.items() if k}
    for key in candidates:
        if key in lower_map and lower_map[key] not in (None, ""):
            return str(lower_map[key]).strip()
    return ""


def parse_csv_export(path: Path) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            title = _pick(row, ("title", "document title", "article title", "ti", "judul"))
            if not title:
                continue
            authors = _pick(row, ("authors", "author", "author(s)", "au", "penulis", "creator"))
            year = _pick(row, ("year", "publication year", "py", "tahun", "date"))[:4]
            venue = _pick(
                row,
                ("source title", "source", "journal", "venue", "publication title", "jo", "t2"),
            )
            doi = normalize_doi(_pick(row, ("doi", "di", "do")))
            url = _pick(row, ("link", "url", "articleurl", "ur", "citesurl"))
            records.append(
                {
                    "authors": authors,
                    "year": year,
                    "title": title,
                    "venue": venue,
                    "doi": doi,
                    "url": url,
                }
            )
    return records


def parse_ris_export(path: Path) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    current: dict[str, Any] = {"authors": []}
    for line in path.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        if line.startswith("ER  -"):
            if current.get("title"):
                records.append(
                    {
                        "authors": "; ".join(current.get("authors", [])),
                        "year": str(current.get("year", ""))[:4],
                        "title": str(current.get("title", "")),
                        "venue": str(current.get("venue", "")),
                        "doi": normalize_doi(current.get("doi", "")),
                        "url": str(current.get("url", "")),
                    }
                )
            current = {"authors": []}
            continue
        m = re.match(r"^([A-Z0-9]{2})\s{2}-\s*(.*)$", line)
        if not m:
            continue
        tag, val = m.group(1), m.group(2).strip()
        if tag in ("AU", "A1"):
            current.setdefault("authors", []).append(val)
        elif tag in ("TI", "T1") and not current.get("title"):
            current["title"] = val
        elif tag in ("PY", "Y1") and not current.get("year"):
            current["year"] = val[:4]
        elif tag in ("JO", "T2", "JF") and not current.get("venue"):
            current["venue"] = val
        elif tag == "DO" and not current.get("doi"):
            current["doi"] = val
        elif tag == "UR" and not current.get("url"):
            current["url"] = val
    return records


def parse_bib_export(path: Path) -> list[dict[str, str]]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    entries = re.split(r"(?m)^@\w+\s*\{", text)[1:]
    records: list[dict[str, str]] = []
    for entry in entries:
        fields: dict[str, str] = {}
        for m in re.finditer(r"(\w+)\s*=\s*(?:\{((?:[^{}]|\{[^{}]*\})*)\}|\"([^\"]*)\"|(\d+))", entry):
            key = m.group(1).lower()
            val = (m.group(2) or m.group(3) or m.group(4) or "").strip()
            fields[key] = re.sub(r"\s+", " ", val)
        if not fields.get("title"):
            continue
        records.append(
            {
                "authors": fields.get("author", ""),
                "year": fields.get("year", "")[:4],
                "title": fields.get("title", ""),
                "venue": fields.get("journal", fields.get("booktitle", "")),
                "doi": normalize_doi(fields.get("doi", "")),
                "url": fields.get("url", ""),
            }
        )
    return records


def ingest_and_deduplicate_exports(
    project_dir: str,
    export_files: list[dict[str, str]],
    reset_existing: bool = True,
) -> dict[str, Any]:
    """Ingest CSV/RIS/BibTeX exports across 5 SPAR-4-SLR pathways, deduplicate, and update Search_Log.xlsx.

    Each item in `export_files` should be:
        {"path": "...", "pathway": "pathway_1_primary_db", "source_db": "Scopus"}
    """
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()
    log_path = cfg.resolve_path(root, "search_log_file")

    seen_dois: set[str] = set()
    seen_title_keys: set[tuple[str, str]] = set()
    pathway_raw: dict[str, int] = {k: 0 for k in PATHWAY_KEYS}
    pathway_unique: dict[str, int] = {k: 0 for k in PATHWAY_KEYS}
    unique_records: list[dict[str, str]] = []
    duplicates_removed = 0

    for spec in export_files:
        fp = Path(spec["path"])
        if not fp.is_absolute():
            fp = root / fp
        if not fp.exists():
            raise FileNotFoundError(f"Export file not found: {fp}")
        pathway = spec.get("pathway", "pathway_1_primary_db")
        if pathway not in PATHWAY_KEYS:
            raise ValueError(f"Invalid pathway '{pathway}'. Must be one of {PATHWAY_KEYS}.")
        source_db = spec.get("source_db", fp.stem)

        ext = fp.suffix.lower()
        if ext == ".csv":
            rows = parse_csv_export(fp)
        elif ext == ".ris":
            rows = parse_ris_export(fp)
        elif ext == ".bib":
            rows = parse_bib_export(fp)
        else:
            raise ValueError(f"Unsupported export extension '{ext}' for {fp.name}. Use .csv, .ris, or .bib.")

        pathway_raw[pathway] += len(rows)
        for rec in rows:
            doi = normalize_doi(rec.get("doi"))
            tkey = (normalize_title(rec.get("title")), str(rec.get("year", ""))[:4])
            is_dup = (doi and doi in seen_dois) or (tkey[0] and tkey in seen_title_keys)
            if is_dup:
                duplicates_removed += 1
                continue
            if doi:
                seen_dois.add(doi)
            if tkey[0]:
                seen_title_keys.add(tkey)
            pathway_unique[pathway] += 1
            unique_records.append(
                {
                    "pathway": pathway,
                    "source_db": source_db,
                    **rec,
                    "doi": doi,
                }
            )

    wb = openpyxl.load_workbook(log_path)
    ws_cand = wb[preset.sheet_candidates]
    if reset_existing and ws_cand.max_row > 1:
        ws_cand.delete_rows(2, ws_cand.max_row - 1)

    start_idx = ws_cand.max_row
    for offset, rec in enumerate(unique_records, start=1):
        rec_id = f"R{start_idx - 1 + offset:04d}"
        ws_cand.append(
            [
                rec_id,
                rec["pathway"],
                rec["source_db"],
                rec["authors"],
                rec["year"],
                rec["title"],
                rec["venue"],
                rec["doi"],
                rec["url"],
                "Pending",
                "Not_Retrieved",
                "Pending",
                "",
                "",
            ]
        )
    style_sheet(ws_cand)

    # Update Flow sheet Assembling counts
    ws_flow = wb[preset.sheet_flow]
    flow_row_map = {ws_flow.cell(r, 1).value: r for r in range(2, ws_flow.max_row + 1)}
    for pkey in PATHWAY_KEYS:
        if pkey in flow_row_map:
            ws_flow.cell(flow_row_map[pkey], 3).value = pathway_unique[pkey]
    total_unique = sum(pathway_unique.values())
    if "total_unique_screened" in flow_row_map:
        ws_flow.cell(flow_row_map["total_unique_screened"], 3).value = total_unique
    style_sheet(ws_flow, {1: 28, 2: 58, 3: 12, 4: 16, 5: 36})
    wb.save(log_path)

    return {
        "status": "ok",
        "files_processed": len(export_files),
        "raw_records_total": sum(pathway_raw.values()),
        "duplicates_removed": duplicates_removed,
        "unique_records_total": total_unique,
        "pathway_raw_counts": pathway_raw,
        "pathway_unique_counts": pathway_unique,
        "search_log_updated": str(log_path),
    }


def reconcile_flow_counts(
    project_dir: str,
    counts_override: dict[str, int] | None = None,
) -> dict[str, Any]:
    """Update (optional) and strictly verify SPAR-4-SLR / PRISMA 2020 5-pathway flow arithmetic."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()
    log_path = cfg.resolve_path(root, "search_log_file")

    wb = openpyxl.load_workbook(log_path)
    ws_flow = wb[preset.sheet_flow]
    row_map = {ws_flow.cell(r, 1).value: r for r in range(2, ws_flow.max_row + 1)}

    if counts_override:
        for k, v in counts_override.items():
            if k not in row_map:
                raise ValueError(f"Unknown flow key '{k}'. Valid keys: {list(row_map)}")
            ws_flow.cell(row_map[k], 3).value = int(v)
        wb.save(log_path)

    vals = {k: int(ws_flow.cell(r, 3).value or 0) for k, r in row_map.items()}
    p_sum = sum(vals.get(k, 0) for k in PATHWAY_KEYS)

    checks = [
        {
            "equation": "assembling_5_pathway_sum",
            "expression": f"{p_sum} == {vals['total_unique_screened']}",
            "passed": p_sum == vals["total_unique_screened"],
        },
        {
            "equation": "title_abstract_screening",
            "expression": (
                f"{vals['total_unique_screened']} - {vals['excluded_title_abstract']} "
                f"== {vals['candidates_passing_ta']}"
            ),
            "passed": (vals["total_unique_screened"] - vals["excluded_title_abstract"])
            == vals["candidates_passing_ta"],
        },
        {
            "equation": "fulltext_retrieval",
            "expression": (
                f"{vals['candidates_passing_ta']} - {vals['unretrieved_fulltext']} "
                f"== {vals['fulltext_assessed']}"
            ),
            "passed": (vals["candidates_passing_ta"] - vals["unretrieved_fulltext"])
            == vals["fulltext_assessed"],
        },
        {
            "equation": "fulltext_eligibility_assessing",
            "expression": (
                f"{vals['fulltext_assessed']} - {vals['excluded_fulltext']} "
                f"== {vals['final_core_corpus']}"
            ),
            "passed": (vals["fulltext_assessed"] - vals["excluded_fulltext"])
            == vals["final_core_corpus"],
        },
    ]

    all_passed = all(c["passed"] for c in checks)
    return {
        "status": "PASS" if all_passed else "FAIL",
        "framework": cfg.framework,
        "counts": vals,
        "checks": checks,
    }
