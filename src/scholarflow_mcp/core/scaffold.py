"""Stage 0: Reproducible academic research project scaffolding (EN & ID presets)."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import openpyxl
from docx import Document

from scholarflow_mcp.config import (
    DEFAULT_CLUSTERS,
    DEFAULT_EXCLUSION_CODES,
    DEFAULT_METHOD_CODES,
    ProjectConfig,
)
from scholarflow_mcp.core.excel_utils import add_safe_range_dropdown, style_sheet

FLOW_KEYS = [
    ("pathway_1_primary_db", "Pathway 1 — Primary database unique records (e.g., Scopus)"),
    ("pathway_2_secondary_db", "Pathway 2 — Secondary / non-primary database unique records"),
    ("pathway_3_web_complement", "Pathway 3 — Complementary web / grey literature records"),
    ("pathway_4_team_seed", "Pathway 4 — Initial team seed corpus (non-duplicate)"),
    ("pathway_5_snowballing", "Pathway 5 — Backward & forward snowballing (non-duplicate)"),
    ("total_unique_screened", "TOTAL UNIQUE RECORDS SCREENED ACROSS ALL PATHWAYS"),
    ("excluded_title_abstract", "Excluded at title/abstract screening or background triage"),
    ("candidates_passing_ta", "Primary study candidates passing title/abstract screening"),
    ("unretrieved_fulltext", "Full text not retrieved (paywall / broken source / unavailable)"),
    ("fulltext_assessed", "Full-text primary studies assessed for eligibility"),
    ("excluded_fulltext", "Excluded at full-text eligibility assessment"),
    ("final_core_corpus", "FINAL LOCKED CORE CORPUS"),
]


def _build_initial_search_log(path: Path, cfg: ProjectConfig) -> None:
    preset = cfg.preset()
    wb = openpyxl.Workbook()

    # Sheet 1: Flow Reconciliation (SPAR-4-SLR / PRISMA 2020)
    ws_flow = wb.active
    ws_flow.title = preset.sheet_flow
    ws_flow.append(["Key", "Stage_Metric", "Count", "SPAR4SLR_Phase", "Notes"])
    phase_map = {
        "pathway_1_primary_db": "Assembling",
        "pathway_2_secondary_db": "Assembling",
        "pathway_3_web_complement": "Assembling",
        "pathway_4_team_seed": "Assembling",
        "pathway_5_snowballing": "Assembling",
        "total_unique_screened": "Assembling",
        "excluded_title_abstract": "Arranging",
        "candidates_passing_ta": "Arranging",
        "unretrieved_fulltext": "Arranging",
        "fulltext_assessed": "Assessing",
        "excluded_fulltext": "Assessing",
        "final_core_corpus": "Assessing",
    }
    for key, label in FLOW_KEYS:
        ws_flow.append([key, label, 0, phase_map[key], ""])
    style_sheet(ws_flow, {1: 28, 2: 58, 3: 12, 4: 16, 5: 36})

    # Sheet 2: Candidates
    ws_cand = wb.create_sheet(preset.sheet_candidates)
    ws_cand.append(
        [
            "Record_ID",
            "Pathway",
            "Source_DB",
            "Authors",
            "Year",
            "Title",
            "Venue",
            "DOI",
            "URL",
            "TA_Decision",
            "Fulltext_Status",
            "FT_Decision",
            "Exclusion_Code",
            "Notes",
        ]
    )
    style_sheet(ws_cand)

    # Sheet 3: Fulltext Download Queue
    ws_q = wb.create_sheet(preset.sheet_download_queue)
    ws_q.append(
        [
            "Queue_ID",
            "Record_ID",
            "Pathway",
            "Priority",
            "Authors",
            "Year",
            "Title",
            "Venue",
            "DOI",
            "Suggested_Filename",
            "Direct_Link",
            "Status",
            "Notes",
        ]
    )
    style_sheet(ws_q)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _build_initial_extraction_matrix(path: Path, cfg: ProjectConfig) -> None:
    preset = cfg.preset()
    wb = openpyxl.Workbook()

    # Sheet 1: Extraction Matrix
    ws_m = wb.active
    ws_m.title = preset.sheet_matrix
    headers = [
        "Study_Code",
        "Authors",
        "Year",
        "Title",
        "Country_Context",
        "Venue",
        "DOI",
        "Paradigm_Cluster",
        "Method_Code",
        "Sample_or_Object",
        "Key_Findings",
        "Reading_Note_Path",
        "PDF_Filename",
    ]
    if cfg.include_ado_tcm:
        headers.extend(
            [
                "ADO_Antecedents",
                "ADO_Decisions",
                "ADO_Outcomes",
                "TCM_Theories",
                "TCM_Contexts",
                "TCM_Methods_Detail",
            ]
        )
    ws_m.append(headers)
    style_sheet(ws_m)

    # Sheet 2: Excluded Full-Text
    ws_ex = wb.create_sheet(preset.sheet_excluded_ft)
    ws_ex.append(
        [
            "Record_ID",
            "Authors",
            "Year",
            "Title",
            "DOI",
            "Exclusion_Code",
            "Reason_Detail",
            "Category",
        ]
    )
    style_sheet(ws_ex)

    # Sheet 3: Codebook (for safe cell-range dropdowns)
    ws_cb = wb.create_sheet(preset.sheet_codebook)
    ws_cb.append(["Paradigm_Clusters", "Method_Codes", "Exclusion_Codes"])
    max_len = max(len(cfg.clusters), len(cfg.method_codes), len(cfg.exclusion_codes))
    for i in range(max_len):
        ws_cb.append(
            [
                cfg.clusters[i] if i < len(cfg.clusters) else "",
                cfg.method_codes[i] if i < len(cfg.method_codes) else "",
                cfg.exclusion_codes[i] if i < len(cfg.exclusion_codes) else "",
            ]
        )
    style_sheet(ws_cb, {1: 42, 2: 42, 3: 44})

    # Sheet 4: Coder 2 Worksheet
    ws_c2 = wb.create_sheet(preset.sheet_coder2)
    ws_c2.append(
        [
            "Study_Code",
            "Authors",
            "Year",
            "Title",
            "Assigned_Coder",
            "Coder2_Paradigm_Cluster",
            "Coder2_Method_Code",
            "Agreement_Status",
            "Coder2_Notes",
        ]
    )
    style_sheet(ws_c2, {1: 14, 2: 26, 3: 8, 4: 44, 5: 20, 6: 36, 7: 34, 8: 20, 9: 32})

    # Attach Excel-safe range-based DataValidation (never >255-char inline lists)
    if cfg.clusters:
        add_safe_range_dropdown(
            ws_c2, "F", 2, 200, preset.sheet_codebook, "A", 2, 1 + len(cfg.clusters),
            prompt_title="Paradigm Cluster",
        )
    if cfg.method_codes:
        add_safe_range_dropdown(
            ws_c2, "G", 2, 200, preset.sheet_codebook, "B", 2, 1 + len(cfg.method_codes),
            prompt_title="Closed Method Code",
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _build_initial_protocol_docx(path: Path, cfg: ProjectConfig) -> None:
    doc = Document()
    title_prefix = "Review Protocol" if cfg.language == "en" else "Protokol Tinjauan Literatur"
    doc.add_heading(f"{title_prefix}: {cfg.project_title}", level=0)
    doc.add_paragraph(f"Authors: {', '.join(cfg.authors)}")
    doc.add_paragraph(f"Framework: {cfg.framework.upper()} | Date: {date.today().isoformat()}")
    doc.add_paragraph("Locked Core Corpus: N = 0 studies (pending screening completion).")

    doc.add_heading("1. Research Questions", level=1)
    rqs = cfg.research_questions or ["RQ1: Define primary research question here."]
    for rq in rqs:
        doc.add_paragraph(rq, style="List Bullet")

    doc.add_heading("2. Closed Coding Taxonomy", level=1)
    doc.add_paragraph("Paradigm Clusters: " + "; ".join(cfg.clusters))
    doc.add_paragraph("Method Codes: " + "; ".join(cfg.method_codes))
    doc.add_paragraph("Exclusion Criteria: " + "; ".join(cfg.exclusion_codes))

    doc.add_heading("3. Integrity & Reproducibility Contract", level=1)
    doc.add_paragraph(
        "All screening counts, physical PDFs, reading notes, extraction codes, bibliographic "
        "exports (.bib/.ris), and manuscript tables must pass the 4-way deterministic audit "
        "before submission."
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def init_research_project(
    project_dir: str,
    project_title: str,
    authors: list[str] | None = None,
    language: str = "en",
    framework: str = "spar4slr",
    research_questions: list[str] | None = None,
    clusters: list[str] | None = None,
    method_codes: list[str] | None = None,
    exclusion_codes: list[str] | None = None,
    include_ado_tcm: bool = True,
) -> dict[str, Any]:
    """Scaffold a complete reproducible academic research / SLR project."""
    lang = language.lower().strip()
    if lang not in ("en", "id"):
        raise ValueError(f"Unsupported language '{language}'. Use 'en' or 'id'.")
    fw = framework.lower().strip()
    if fw not in ("spar4slr", "prisma2020", "empirical"):
        raise ValueError(f"Unsupported framework '{framework}'. Use 'spar4slr', 'prisma2020', or 'empirical'.")

    root = Path(project_dir).resolve()
    root.mkdir(parents=True, exist_ok=True)

    cfg = ProjectConfig(
        project_title=project_title,
        authors=authors or ["Lead Researcher"],
        language=lang,
        framework=fw,
        research_questions=research_questions or [],
        clusters=clusters or list(DEFAULT_CLUSTERS),
        method_codes=method_codes or list(DEFAULT_METHOD_CODES),
        exclusion_codes=exclusion_codes or list(DEFAULT_EXCLUSION_CODES),
        include_ado_tcm=include_ado_tcm,
    )

    dir_keys = [
        "admin_dir",
        "manuscript_dir",
        "figures_dir",
        "literature_dir",
        "core_dir",
        "unclassified_dir",
        "excluded_dir",
        "methodology_dir",
        "regulation_dir",
        "background_dir",
        "misfiltered_dir",
        "extraction_dir",
        "reading_notes_dir",
        "exports_dir",
        "tools_dir",
    ]
    created_dirs: list[str] = []
    for k in dir_keys:
        d = cfg.resolve_path(root, k)
        d.mkdir(parents=True, exist_ok=True)
        created_dirs.append(str(d.relative_to(root)))

    cfg_path = cfg.save(root)
    search_log = cfg.resolve_path(root, "search_log_file")
    matrix_path = cfg.resolve_path(root, "extraction_matrix_file")
    protocol_path = cfg.resolve_path(root, "protocol_file")

    _build_initial_search_log(search_log, cfg)
    _build_initial_extraction_matrix(matrix_path, cfg)
    _build_initial_protocol_docx(protocol_path, cfg)

    # Create starter excluded log and misfiltered manifest
    excl_md = cfg.resolve_path(root, "excluded_log_md")
    if not excl_md.exists():
        excl_md.write_text("# Full-Text Excluded Studies Log\n\n", encoding="utf-8")

    manifest_csv = cfg.resolve_path(root, "misfiltered_manifest_csv")
    if not manifest_csv.exists():
        manifest_csv.write_text("filename,original_source,reason,archived_at\n", encoding="utf-8")

    readme_path = root / "README.md"
    if not readme_path.exists():
        readme_path.write_text(
            f"# {project_title}\n\n"
            f"- **Authors:** {', '.join(cfg.authors)}\n"
            f"- **Framework:** `{cfg.framework.upper()}`\n"
            f"- **Language Preset:** `{cfg.language}`\n"
            f"- **Managed by:** `scholarflow-mcp`\n",
            encoding="utf-8",
        )

    return {
        "status": "ok",
        "project_root": str(root),
        "config_path": str(cfg_path),
        "language": lang,
        "framework": fw,
        "created_directories": created_dirs,
        "created_files": [
            str(cfg_path.relative_to(root)),
            str(search_log.relative_to(root)),
            str(matrix_path.relative_to(root)),
            str(protocol_path.relative_to(root)),
            str(excl_md.relative_to(root)),
            str(manifest_csv.relative_to(root)),
            "README.md",
        ],
    }
