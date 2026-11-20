"""Stage 8: 4-Way Deterministic Project Integrity Auditor (20+ checks) & Project Status Dashboard."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import openpyxl
from docx import Document

from scholarflow_mcp.config import ProjectConfig
from scholarflow_mcp.core.corpus import check_pdf_readability
from scholarflow_mcp.core.excel_utils import detect_unsafe_inline_validations
from scholarflow_mcp.core.records import reconcile_flow_counts


def run_four_way_integrity_audit(
    project_dir: str,
    save_report_json: bool = True,
) -> dict[str, Any]:
    """Run 20+ deterministic 4-way integrity checks across PDFs, Excel Logs/Matrix, Reading Notes, Manuscript, Bib/RIS & Protocol."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()

    checks: list[dict[str, Any]] = []

    def add_check(check_id: str, category: str, passed: bool, message: str, detail: Any = None) -> None:
        checks.append(
            {
                "id": check_id,
                "category": category,
                "passed": bool(passed),
                "message": message,
                "detail": detail,
            }
        )

    # =========================================================================
    # DIRECTION 1: Physical Literature Folders & PDF Readability
    # =========================================================================
    unclass_dir = cfg.resolve_path(root, "unclassified_dir")
    core_dir = cfg.resolve_path(root, "core_dir")
    excl_dir = cfg.resolve_path(root, "excluded_dir")
    bg_dir = cfg.resolve_path(root, "background_dir")
    mis_dir = cfg.resolve_path(root, "misfiltered_dir")

    unclass_files = [f.name for f in unclass_dir.iterdir() if f.is_file()] if unclass_dir.exists() else []
    add_check(
        "D1_01_unclassified_empty",
        "Physical_Corpus",
        len(unclass_files) == 0,
        "2_Unclassified staging folder is empty (0 hanging files)"
        if not unclass_files
        else f"Pending unclassified files remain: {len(unclass_files)}",
        unclass_files,
    )

    core_files = sorted(f.name for f in core_dir.iterdir() if f.is_file()) if core_dir.exists() else []
    add_check(
        "D1_02_core_files_present",
        "Physical_Corpus",
        len(core_files) > 0,
        f"1_Core folder contains {len(core_files)} verified study file(s)",
        {"count": len(core_files)},
    )

    corrupt_pdfs: list[dict[str, str]] = []
    pdf_checked = 0
    for folder in (core_dir, excl_dir, bg_dir, mis_dir):
        if not folder.exists():
            continue
        for fp in folder.iterdir():
            if fp.is_file() and fp.suffix.lower() == ".pdf":
                ok, _, err = check_pdf_readability(fp)
                pdf_checked += 1
                if not ok:
                    corrupt_pdfs.append({"file": str(fp.relative_to(root)), "error": err})
    add_check(
        "D1_03_all_pdfs_readable",
        "Physical_Corpus",
        len(corrupt_pdfs) == 0,
        f"All {pdf_checked} PDFs across literature folders open cleanly via PyMuPDF"
        if not corrupt_pdfs
        else f"Found {len(corrupt_pdfs)} unreadable/corrupt PDF(s)",
        corrupt_pdfs,
    )

    # =========================================================================
    # DIRECTION 2: Search Log 5-Pathway Arithmetic & Clickable Download Queue
    # =========================================================================
    flow_res = reconcile_flow_counts(str(root))
    flow_vals = flow_res["counts"]
    for idx, eq in enumerate(flow_res["checks"], start=1):
        add_check(
            f"D2_0{idx}_{eq['equation']}",
            "Search_Log",
            eq["passed"],
            f"Flow equation `{eq['equation']}`: {eq['expression']}",
        )

    log_path = cfg.resolve_path(root, "search_log_file")
    wb_l = openpyxl.load_workbook(log_path, data_only=False)
    ws_q = wb_l[preset.sheet_download_queue]
    q_rows = [r for r in range(2, ws_q.max_row + 1) if ws_q.cell(r, 1).value]
    q_links = sum(1 for r in q_rows if ws_q.cell(r, 11).hyperlink is not None)
    add_check(
        "D2_05_manual_queue_matches_unretrieved",
        "Search_Log",
        len(q_rows) == flow_vals.get("unretrieved_fulltext", 0),
        f"Fulltext_Download_Queue rows ({len(q_rows)}) match unretrieved count ({flow_vals.get('unretrieved_fulltext', 0)})",
    )
    add_check(
        "D2_06_manual_queue_clickable_links",
        "Search_Log",
        len(q_rows) == q_links,
        f"Clickable hyperlinks on manual download queue: {q_links}/{len(q_rows)}",
    )

    # =========================================================================
    # DIRECTION 3: Extraction Matrix, Safe Excel DataValidation & 1-to-1 Reading Notes
    # =========================================================================
    matrix_path = cfg.resolve_path(root, "extraction_matrix_file")
    wb_m = openpyxl.load_workbook(matrix_path, data_only=False)
    ws_m = wb_m[preset.sheet_matrix]
    m_headers = [c.value for c in ws_m[1]]
    col = {h: i for i, h in enumerate(m_headers) if h}
    m_rows = [r for r in list(ws_m.iter_rows(values_only=True))[1:] if r and r[0]]
    m_codes = [str(r[col["Study_Code"]]).strip() for r in m_rows]

    add_check(
        "D3_01_matrix_unique_codes_match_flow",
        "Extraction_Matrix",
        len(m_codes) == len(set(m_codes)) == flow_vals.get("final_core_corpus", 0) and len(m_codes) > 0,
        f"Extraction_Matrix has {len(m_codes)} unique study codes matching locked flow N={flow_vals.get('final_core_corpus', 0)}",
    )

    core_prefixes = sorted(f.split("_")[0] for f in core_files)
    add_check(
        "D3_02_matrix_codes_match_core_files_1to1",
        "Extraction_Matrix",
        sorted(m_codes) == core_prefixes,
        "1-to-1 match between Extraction_Matrix study codes and 1_Core physical files"
        if sorted(m_codes) == core_prefixes
        else f"Code/file mismatch: {sorted(set(m_codes) ^ set(core_prefixes))}",
        sorted(set(m_codes) ^ set(core_prefixes)),
    )

    missing_taxonomy = [
        str(r[col["Study_Code"]])
        for r in m_rows
        if not str(r[col.get("Paradigm_Cluster", 7)] or "").strip()
        or not str(r[col.get("Method_Code", 8)] or "").strip()
    ]
    add_check(
        "D3_03_closed_taxonomy_complete",
        "Extraction_Matrix",
        len(missing_taxonomy) == 0 and len(m_rows) > 0,
        "Paradigm_Cluster & Method_Code populated for 100% of core studies"
        if not missing_taxonomy
        else f"Missing closed taxonomy labels on: {missing_taxonomy}",
        missing_taxonomy,
    )

    extraction_dir = cfg.resolve_path(root, "extraction_dir")
    missing_notes: list[str] = []
    for r in m_rows:
        rel_note = str(r[col.get("Reading_Note_Path", 11)] or "").strip()
        if not rel_note or not (extraction_dir / rel_note).exists():
            missing_notes.append(f"{r[col['Study_Code']]} ({rel_note or 'empty'})")
    add_check(
        "D3_04_reading_notes_1to1",
        "Extraction_Matrix",
        len(missing_notes) == 0 and len(m_rows) > 0,
        f"All {len(m_rows)} core studies have existing 1-to-1 Reading_Notes Markdown files"
        if not missing_notes
        else f"Missing reading notes: {missing_notes}",
        missing_notes,
    )

    ws_c2 = wb_m[preset.sheet_coder2]
    c2_codes = [str(r[0]).strip() for r in list(ws_c2.iter_rows(values_only=True))[1:] if r and r[0]]
    add_check(
        "D3_05_coder2_sheet_synced",
        "Extraction_Matrix",
        c2_codes == m_codes,
        f"Coder2_Worksheet synced row-for-row with Extraction_Matrix ({len(c2_codes)}/{len(m_codes)})",
    )

    unsafe_dvs = detect_unsafe_inline_validations(ws_m) + detect_unsafe_inline_validations(ws_c2)
    add_check(
        "D3_06_excel_datavalidation_safe",
        "Extraction_Matrix",
        len(unsafe_dvs) == 0,
        "All Excel DataValidations use safe cell-range formulas (0 inline >255-char corruption risks)",
        unsafe_dvs,
    )

    ws_ex = wb_m[preset.sheet_excluded_ft]
    ex_rows = [
        r
        for r in list(ws_ex.iter_rows(values_only=True))[1:]
        if r and r[0] and str(r[7] or "Excluded").strip().lower() in {"excluded", "dikeluarkan"}
    ]
    add_check(
        "D3_07_excluded_ft_matches_flow",
        "Extraction_Matrix",
        len(ex_rows) == flow_vals.get("excluded_fulltext", 0),
        f"Excluded_FullText rows ({len(ex_rows)}) match Search_Log excluded_fulltext ({flow_vals.get('excluded_fulltext', 0)})",
    )

    # =========================================================================
    # DIRECTION 4: Manuscript DOCX, Figures, Citations, Bib/RIS & Protocol
    # =========================================================================
    docx_path = cfg.resolve_path(root, "manuscript_docx_file")
    if docx_path.exists():
        doc = Document(docx_path)
        paras = [p.text for p in doc.paragraphs]
        table_texts = [" | ".join(c.text for c in row.cells) for t in doc.tables for row in t.rows]
        full_text = "\n".join(paras + table_texts)
        full_paras = "\n".join(paras)

        # Abstract word count
        abs_words = 0
        try:
            a_start = next(i for i, p in enumerate(paras) if p.strip().upper() in {"ABSTRACT", "ABSTRAK"})
            a_end = next(
                i for i in range(a_start + 1, len(paras)) if paras[i].strip().lower().startswith(("keywords:", "kata kunci:"))
            )
            abs_words = len(" ".join(paras[a_start + 1 : a_end]).split())
        except StopIteration:
            abs_words = 0
        add_check(
            "D4_01_abstract_word_count",
            "Manuscript_Bib_Protocol",
            cfg.abstract_min_words <= abs_words <= cfg.abstract_max_words,
            f"Structured abstract word count = {abs_words} (target {cfg.abstract_min_words}–{cfg.abstract_max_words})",
        )

        # Embedded figures & explicit text reference
        n_inline = len(doc.inline_shapes)
        add_check(
            "D4_02_figures_embedded",
            "Manuscript_Bib_Protocol",
            n_inline >= cfg.expected_figures,
            f"Manuscript embeds {n_inline} figure(s) (expected >= {cfg.expected_figures})",
        )
        fig_refs_ok = all(
            full_paras.count(f"Figure {i}") >= 2 or full_paras.count(f"Gambar {i}") >= 2
            for i in range(1, cfg.expected_figures + 1)
        )
        add_check(
            "D4_03_figures_referenced_in_prose",
            "Manuscript_Bib_Protocol",
            fig_refs_ok,
            f"All {cfg.expected_figures} figures explicitly referenced in prose and captioned",
        )

        # Appendix table row-for-row sync with Extraction_Matrix
        app_codes: list[str] = []
        if doc.tables:
            last_tbl = doc.tables[-1]
            app_codes = [row.cells[1].text.strip() for row in last_tbl.rows[1:] if len(row.cells) > 1]
        add_check(
            "D4_04_appendix_table_synced",
            "Manuscript_Bib_Protocol",
            app_codes == m_codes,
            f"Manuscript appendix table ({len(app_codes)} rows) matches Extraction_Matrix row-for-row",
        )

        # Forbidden legacy patterns
        found_forbidden = [pat for pat in cfg.forbidden_patterns if pat.lower() in full_text.lower()]
        add_check(
            "D4_05_zero_forbidden_patterns",
            "Manuscript_Bib_Protocol",
            len(found_forbidden) == 0,
            "Manuscript is 100% free of forbidden placeholders / legacy citation errors"
            if not found_forbidden
            else f"Forbidden patterns found: {found_forbidden}",
            found_forbidden,
        )

        # All study codes cited explicitly in manuscript
        missing_codes = [c for c in m_codes if f"[{c}]" not in full_text]
        add_check(
            "D4_06_all_study_codes_mapped",
            "Manuscript_Bib_Protocol",
            len(missing_codes) == 0 and len(m_codes) > 0,
            f"All {len(m_codes)} study codes [{m_codes[0] if m_codes else ''}].. mapped in manuscript"
            if not missing_codes
            else f"Study codes absent from manuscript: {missing_codes}",
            missing_codes,
        )

        # 100% of Reference list cited in body prose
        uncited_refs: list[str] = []
        ref_count = 0
        try:
            r_start = next(i for i, p in enumerate(paras) if p.strip().upper() in {"REFERENCES", "DAFTAR PUSTAKA"})
            r_end = next(
                (i for i in range(r_start + 1, len(paras)) if paras[i].strip().upper().startswith(("APPENDIX", "LAMPIRAN"))),
                len(paras),
            )
            refs = [p.strip() for p in paras[r_start + 1 : r_end] if p.strip()]
            ref_count = len(refs)
            body_prose = "\n".join(paras[:r_start]).lower()
            for ref in refs:
                m_ref = re.match(r"^([^,(]+?)(?:,|\s*\().*?\((\d{4})", ref)
                if not m_ref:
                    continue
                surname = m_ref.group(1).strip().rstrip(".").split()[-1].lower()
                yr = m_ref.group(2)
                if not (surname in body_prose and yr in body_prose):
                    uncited_refs.append(f"{surname} ({yr})")
        except StopIteration:
            uncited_refs.append("REFERENCES section missing")
        add_check(
            "D4_07_references_100pct_cited_in_prose",
            "Manuscript_Bib_Protocol",
            len(uncited_refs) == 0 and ref_count >= len(m_codes) and ref_count > 0,
            f"All {ref_count}/{ref_count} reference list entries are cited in main body prose"
            if not uncited_refs
            else f"Uncited references in prose: {uncited_refs}",
            uncited_refs,
        )
    else:
        for cid in (
            "D4_01_abstract_word_count",
            "D4_02_figures_embedded",
            "D4_03_figures_referenced_in_prose",
            "D4_04_appendix_table_synced",
            "D4_05_zero_forbidden_patterns",
            "D4_06_all_study_codes_mapped",
            "D4_07_references_100pct_cited_in_prose",
        ):
            add_check(cid, "Manuscript_Bib_Protocol", False, f"Manuscript DOCX not found: {docx_path.name}")
        ref_count = 0

    # BibTeX & RIS validation (including zero '&' author bug)
    bib_path = cfg.resolve_path(root, "bib_file")
    ris_path = cfg.resolve_path(root, "ris_file")
    if bib_path.exists() and ris_path.exists():
        ris_lines = ris_path.read_text(encoding="utf-8", errors="replace").splitlines()
        bad_au = [ln for ln in ris_lines if ln.startswith("AU  -") and "&" in ln]
        ris_count = sum(1 for ln in ris_lines if ln.startswith("ER  -"))
        bib_text = bib_path.read_text(encoding="utf-8", errors="replace")
        bib_count = len(re.findall(r"(?m)^@\w+\s*\{", bib_text))
        add_check(
            "D4_08_bib_and_ris_clean",
            "Manuscript_Bib_Protocol",
            bib_count == ris_count >= len(m_codes) and ris_count > 0 and len(bad_au) == 0,
            f"BibTeX ({bib_count}) & RIS ({ris_count}) valid with 0 malformed '&' AU lines"
            if not bad_au
            else f"RIS/Bib mismatch or malformed AU lines: {bad_au[:3]}",
            {"bib_count": bib_count, "ris_count": ris_count, "bad_au": bad_au},
        )
    else:
        add_check("D4_08_bib_and_ris_clean", "Manuscript_Bib_Protocol", False, "Missing .bib or .ris export file")

    # Protocol sync check
    prot_path = cfg.resolve_path(root, "protocol_file")
    if prot_path.exists():
        pdoc = Document(prot_path)
        ptext = "\n".join(p.text for p in pdoc.paragraphs)
        n_token = f"N = {len(m_codes)}"
        prot_ok = len(ptext) > 200 and n_token in ptext
        add_check(
            "D4_09_protocol_synced_with_corpus",
            "Manuscript_Bib_Protocol",
            prot_ok,
            f"Protocol document valid and synchronized with locked corpus ({n_token})"
            if prot_ok
            else f"Protocol missing locked corpus token '{n_token}'",
        )
    else:
        add_check("D4_09_protocol_synced_with_corpus", "Manuscript_Bib_Protocol", False, "Protocol DOCX missing")

    passed_count = sum(1 for c in checks if c["passed"])
    failed_checks = [c for c in checks if not c["passed"]]
    overall = "PASS" if not failed_checks else "FAIL"

    report = {
        "status": overall,
        "project_title": cfg.project_title,
        "framework": cfg.framework,
        "passed": passed_count,
        "failed": len(failed_checks),
        "total_checks": len(checks),
        "checks": checks,
        "failed_checks": failed_checks,
    }

    if save_report_json:
        out_json = cfg.resolve_path(root, "tools_dir") / "integrity_audit_report.json"
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        report["report_path"] = str(out_json)

    return report


def get_project_status(project_dir: str) -> dict[str, Any]:
    """Return a concise dashboard of the project's current state and recommended next step."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()

    core_dir = cfg.resolve_path(root, "core_dir")
    unclass_dir = cfg.resolve_path(root, "unclassified_dir")
    notes_dir = cfg.resolve_path(root, "reading_notes_dir")
    audit_json = cfg.resolve_path(root, "tools_dir") / "integrity_audit_report.json"

    n_core_files = len([f for f in core_dir.iterdir() if f.is_file()]) if core_dir.exists() else 0
    n_unclass = len([f for f in unclass_dir.iterdir() if f.is_file()]) if unclass_dir.exists() else 0
    n_notes = (
        len([f for f in notes_dir.iterdir() if f.is_file() and not f.name.startswith("_")])
        if notes_dir.exists()
        else 0
    )

    wb_m = openpyxl.load_workbook(cfg.resolve_path(root, "extraction_matrix_file"), data_only=True)
    ws_m = wb_m[preset.sheet_matrix]
    n_matrix = sum(1 for r in range(2, ws_m.max_row + 1) if ws_m.cell(r, 1).value)

    flow_res = reconcile_flow_counts(str(root))
    last_audit = None
    if audit_json.exists():
        try:
            a_obj = json.loads(audit_json.read_text(encoding="utf-8"))
            last_audit = f"{a_obj.get('status')} ({a_obj.get('passed')}/{a_obj.get('total_checks')})"
        except Exception:
            last_audit = "unreadable"

    if flow_res["counts"].get("total_unique_screened", 0) == 0:
        next_action = "Run `ingest_and_deduplicate_exports` to import multi-database exports."
    elif n_unclass > 0:
        next_action = f"Route {n_unclass} pending file(s) from 2_Unclassified via `route_corpus_files`."
    elif n_matrix == 0:
        next_action = "Populate core studies and closed taxonomy via `configure_extraction_matrix`."
    elif not cfg.resolve_path(root, "bib_file").exists():
        next_action = "Export clean BibTeX/RIS via `verify_and_export_bibliography`."
    elif not cfg.resolve_path(root, "manuscript_docx_file").exists():
        next_action = "Generate 300 DPI figures and sync manuscript via `sync_manuscript_snapshot`."
    elif last_audit is None or not last_audit.startswith("PASS"):
        next_action = "Run `run_four_way_integrity_audit` and fix any FAIL items before submission."
    else:
        next_action = "Project passed 4-way integrity audit — ready for peer-review simulation / submission!"

    return {
        "project_title": cfg.project_title,
        "language": cfg.language,
        "framework": cfg.framework,
        "counts": {
            "total_unique_screened": flow_res["counts"].get("total_unique_screened", 0),
            "core_files": n_core_files,
            "unclassified_files": n_unclass,
            "matrix_rows": n_matrix,
            "reading_notes": n_notes,
            "flow_reconciliation_status": flow_res["status"],
        },
        "last_audit": last_audit or "not_run",
        "recommended_next_action": next_action,
    }
