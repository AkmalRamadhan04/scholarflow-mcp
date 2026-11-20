"""Stage 3: Non-destructive physical corpus routing, PyMuPDF readability check & 1-to-1 reading notes."""

from __future__ import annotations

import csv
import os
import re
import shutil
import stat
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pymupdf

from scholarflow_mcp.config import ProjectConfig

VALID_DESTINATIONS = {
    "core": "core_dir",
    "excluded": "excluded_dir",
    "background": "background_dir",
    "methodology": "methodology_dir",
    "regulation": "regulation_dir",
    "misfiltered": "misfiltered_dir",
}


def check_pdf_readability(pdf_path: Path) -> tuple[bool, int, str]:
    """Verify that a PDF opens cleanly via PyMuPDF and has >= 1 page."""
    try:
        doc = pymupdf.open(pdf_path)
        pages = len(doc)
        doc.close()
        if pages < 1:
            return False, 0, "PDF has 0 pages"
        return True, pages, "ok"
    except Exception as exc:
        return False, 0, f"{type(exc).__name__}: {exc}"


def _safe_move(src: Path, dst: Path) -> str:
    if not src.exists():
        if dst.exists():
            return "already_moved"
        raise FileNotFoundError(f"Source file not found: {src}")
    dst.parent.mkdir(parents=True, exist_ok=True)
    os.chmod(src, os.stat(src).st_mode | stat.S_IWUSR | stat.S_IRUSR)
    shutil.move(str(src), str(dst))
    return "moved"


def _sanitize_slug(text: str, max_len: int = 24) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "", text or "")
    return cleaned[:max_len] or "Study"


def _render_reading_note(item: dict[str, Any], language: str, include_ado_tcm: bool) -> str:
    code = item.get("study_code", "S01")
    authors = item.get("authors", "Unknown")
    year = item.get("year", "n.d.")
    title = item.get("title", "Untitled")
    doi = item.get("doi", "-")
    cluster = item.get("paradigm_cluster", "Pending")
    method_code = item.get("method_code", "Pending")
    findings = item.get("key_findings", "_To be completed after full-text extraction._")

    if language == "id":
        lines = [
            f"# [{code}] {authors} ({year}) — {title}",
            "",
            f"- **Kode Studi:** `{code}`",
            f"- **DOI / Tautan:** {doi}",
            f"- **Klaster Paradigma:** {cluster}",
            f"- **Kode Tertutup Metode:** {method_code}",
            "",
            "## Ringkasan Metodologi & Temuan Utama",
            str(findings),
            "",
        ]
        if include_ado_tcm:
            lines.extend(
                [
                    "## Ekstraksi ADO–TCM",
                    f"- **Antecedents (A):** {item.get('ado_antecedents', '-')}",
                    f"- **Decisions (D):** {item.get('ado_decisions', '-')}",
                    f"- **Outcomes (O):** {item.get('ado_outcomes', '-')}",
                    f"- **Theories (T):** {item.get('tcm_theories', '-')}",
                    f"- **Contexts (C):** {item.get('tcm_contexts', '-')}",
                    f"- **Methods (M):** {item.get('tcm_methods_detail', '-')}",
                    "",
                ]
            )
    else:
        lines = [
            f"# [{code}] {authors} ({year}) — {title}",
            "",
            f"- **Study Code:** `{code}`",
            f"- **DOI / Link:** {doi}",
            f"- **Paradigm Cluster:** {cluster}",
            f"- **Closed Method Code:** {method_code}",
            "",
            "## Methodology & Key Findings",
            str(findings),
            "",
        ]
        if include_ado_tcm:
            lines.extend(
                [
                    "## ADO–TCM Extraction",
                    f"- **Antecedents (A):** {item.get('ado_antecedents', '-')}",
                    f"- **Decisions (D):** {item.get('ado_decisions', '-')}",
                    f"- **Outcomes (O):** {item.get('ado_outcomes', '-')}",
                    f"- **Theories (T):** {item.get('tcm_theories', '-')}",
                    f"- **Contexts (C):** {item.get('tcm_contexts', '-')}",
                    f"- **Methods (M):** {item.get('tcm_methods_detail', '-')}",
                    "",
                ]
            )
    return "\n".join(lines)


def route_corpus_files(
    project_dir: str,
    decisions: list[dict[str, Any]],
) -> dict[str, Any]:
    """Non-destructively route files from `2_Unclassified` into core/excluded/background/misfiltered.

    Each item in `decisions` has:
        - `source_filename`: filename inside `2_Unclassified` (or absolute path)
        - `destination`: one of `core`, `excluded`, `background`, `methodology`, `regulation`, `misfiltered`
        - `study_code`: required for `core` (e.g., `"S01"`)
        - `target_filename`: optional explicit target filename; auto-generated if omitted
        - `authors`, `year`, `title`, `doi`, `exclusion_code`, `reason`, `paradigm_cluster`, `method_code`, etc.
    """
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    unclass_dir = cfg.resolve_path(root, "unclassified_dir")
    notes_dir = cfg.resolve_path(root, "reading_notes_dir")
    excl_md = cfg.resolve_path(root, "excluded_log_md")
    manifest_csv = cfg.resolve_path(root, "misfiltered_manifest_csv")

    routed: list[dict[str, Any]] = []
    corrupt_files: list[dict[str, str]] = []
    notes_created: list[str] = []

    for item in decisions:
        dest_key = str(item.get("destination", "")).lower().strip()
        if dest_key not in VALID_DESTINATIONS:
            raise ValueError(
                f"Invalid destination '{dest_key}'. Choose from {sorted(VALID_DESTINATIONS)}."
            )
        src_raw = str(item["source_filename"])
        src_path = Path(src_raw) if Path(src_raw).is_absolute() else (unclass_dir / src_raw)

        # Validate PDF readability before routing into core
        if src_path.exists() and src_path.suffix.lower() == ".pdf":
            ok, pages, msg = check_pdf_readability(src_path)
            if not ok:
                corrupt_files.append({"file": str(src_path), "error": msg})
                if dest_key == "core":
                    continue
        else:
            pages = 0

        dest_dir = cfg.resolve_path(root, VALID_DESTINATIONS[dest_key])
        study_code = str(item.get("study_code", "")).strip()
        if dest_key == "core" and not study_code:
            raise ValueError(f"`study_code` is required when routing '{src_path.name}' to 'core'.")

        if item.get("target_filename"):
            target_name = str(item["target_filename"])
        elif study_code:
            first_auth = str(item.get("authors", "Author")).split(";")[0].split(",")[0].strip()
            yr = str(item.get("year", "nd"))[:4]
            ext = src_path.suffix or ".pdf"
            target_name = f"{study_code}_{_sanitize_slug(first_auth)}_{yr}{ext}"
        else:
            target_name = src_path.name

        dst_path = dest_dir / target_name
        move_status = _safe_move(src_path, dst_path)

        note_rel = ""
        if dest_key == "core":
            first_auth = str(item.get("authors", "Author")).split(";")[0].split(",")[0].strip()
            yr = str(item.get("year", "nd"))[:4]
            note_name = f"{study_code}_{_sanitize_slug(first_auth)}_{yr}.md"
            note_path = notes_dir / note_name
            note_path.parent.mkdir(parents=True, exist_ok=True)
            note_path.write_text(
                _render_reading_note(item, cfg.language, cfg.include_ado_tcm),
                encoding="utf-8",
            )
            note_rel = str(note_path.relative_to(cfg.resolve_path(root, "extraction_dir"))).replace("\\", "/")
            notes_created.append(note_rel)
        elif dest_key == "excluded":
            excl_code = item.get("exclusion_code", "E1")
            reason = item.get("reason", "Excluded after full-text review.")
            entry = (
                f"- **{study_code or target_name}** — {item.get('authors', '')} ({item.get('year', '')}): "
                f"*{item.get('title', target_name)}* | **Code:** `{excl_code}` | **Reason:** {reason}\n"
            )
            with excl_md.open("a", encoding="utf-8") as f:
                f.write(entry)
        elif dest_key == "misfiltered":
            manifest_csv.parent.mkdir(parents=True, exist_ok=True)
            write_header = not manifest_csv.exists()
            with manifest_csv.open("a", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                if write_header:
                    writer.writerow(["filename", "original_source", "reason", "archived_at"])
                writer.writerow(
                    [
                        target_name,
                        item.get("original_source", "search_import"),
                        item.get("reason", "Out-of-topic misfiltered query result"),
                        datetime.now(UTC).isoformat(timespec="seconds"),
                    ]
                )

        routed.append(
            {
                "study_code": study_code,
                "destination": dest_key,
                "status": move_status,
                "target_path": str(dst_path.relative_to(root)),
                "pdf_pages": pages,
                "reading_note_rel": note_rel,
            }
        )

    return {
        "status": "ok" if not corrupt_files else "partial_with_corrupt_pdfs",
        "routed_count": len(routed),
        "notes_created_count": len(notes_created),
        "corrupt_pdfs_count": len(corrupt_files),
        "routed": routed,
        "corrupt_pdfs": corrupt_files,
    }
