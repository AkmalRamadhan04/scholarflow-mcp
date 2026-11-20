"""FastMCP server exposing all 14 ScholarFlow tools, 3 resources, and 3 guided prompts."""

from __future__ import annotations

import json
from typing import Any

try:
    from mcp.server.mcpserver import MCPServer
except ImportError:  # pragma: no cover — compatibility fallback for mcp<2
    from mcp.server.fastmcp import FastMCP as MCPServer  # type: ignore[no-redef]

from scholarflow_mcp.core import (
    acquisition,
    audit,
    bibliography,
    coding,
    corpus,
    records,
    reliability,
    scaffold,
    sync_report,
)

mcp = MCPServer(
    "scholarflow-mcp",
    instructions=(
        "Unified Academic Research & Reproducible SLR MCP Server. "
        "Provides deterministic tools for SPAR-4-SLR / PRISMA 2020 5-pathway reconciliation, "
        "legal OA PDF retrieval + clickable manual download queues, non-destructive corpus routing, "
        "closed-taxonomy ADO-TCM extraction matrices with Excel-safe dropdowns, 100%-blinded human "
        "second-coder exports & Cohen's Kappa/PABAK reliability, BibTeX/RIS verification, "
        "300 DPI data-driven figures/manuscript sync, and 4-way deterministic integrity auditing."
    ),
)


def _as_json(payload: dict[str, Any]) -> str:
    return json.dumps(payload, indent=2, ensure_ascii=False)


# =============================================================================
# 14 DETERMINISTIC MCP TOOLS
# =============================================================================


@mcp.tool()
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
) -> str:
    """Stage 0: Initialize a reproducible academic research or SLR workspace with EN/ID presets, Excel logs, and protocol."""
    return _as_json(
        scaffold.init_research_project(
            project_dir=project_dir,
            project_title=project_title,
            authors=authors,
            language=language,
            framework=framework,
            research_questions=research_questions,
            clusters=clusters,
            method_codes=method_codes,
            exclusion_codes=exclusion_codes,
            include_ado_tcm=include_ado_tcm,
        )
    )


@mcp.tool()
def ingest_and_deduplicate_exports(
    project_dir: str,
    export_files: list[dict[str, str]],
    reset_existing: bool = True,
) -> str:
    """Stage 1a: Parse CSV/RIS/BibTeX exports across 5 SPAR-4-SLR pathways, deduplicate by DOI/title, and update Search_Log.xlsx."""
    return _as_json(
        records.ingest_and_deduplicate_exports(
            project_dir=project_dir,
            export_files=export_files,
            reset_existing=reset_existing,
        )
    )


@mcp.tool()
def reconcile_flow_counts(
    project_dir: str,
    counts_override: dict[str, int] | None = None,
) -> str:
    """Stage 1b: Update and verify strict 5-pathway SPAR-4-SLR / PRISMA 2020 flow accounting equations."""
    return _as_json(
        records.reconcile_flow_counts(
            project_dir=project_dir,
            counts_override=counts_override,
        )
    )


@mcp.tool()
def fetch_open_access_pdfs(
    project_dir: str,
    candidates: list[dict[str, str]],
    timeout_seconds: float = 15.0,
) -> str:
    """Stage 2a: Legally retrieve OA PDFs via OpenAlex API & citation_pdf_url/OJS metadata (strictly refuses CAPTCHAs; verifies via PyMuPDF)."""
    return _as_json(
        acquisition.fetch_open_access_pdfs(
            project_dir=project_dir,
            candidates=candidates,
            timeout_seconds=timeout_seconds,
        )
    )


@mcp.tool()
def build_manual_download_queue(
    project_dir: str,
    queue_items: list[dict[str, str]] | None = None,
) -> str:
    """Stage 2b: Build a prioritized manual download queue with clickable hyperlinks in Search_Log.xlsx."""
    return _as_json(
        acquisition.build_manual_download_queue(
            project_dir=project_dir,
            queue_items=queue_items,
        )
    )


@mcp.tool()
def route_corpus_files(
    project_dir: str,
    decisions: list[dict[str, Any]],
) -> str:
    """Stage 3: Non-destructively route files into core/excluded/background/misfiltered (+ manifest), verify PDF readability, and generate 1-to-1 reading notes."""
    return _as_json(
        corpus.route_corpus_files(
            project_dir=project_dir,
            decisions=decisions,
        )
    )


@mcp.tool()
def configure_extraction_matrix(
    project_dir: str,
    clusters: list[str] | None = None,
    method_codes: list[str] | None = None,
    exclusion_codes: list[str] | None = None,
    studies: list[dict[str, Any]] | None = None,
    excluded_studies: list[dict[str, Any]] | None = None,
    replace_existing_studies: bool = False,
) -> str:
    """Stage 4: Configure closed taxonomy codebooks, Excel-safe cell-range dropdowns, ADO-TCM columns, and upsert core/excluded study rows."""
    return _as_json(
        coding.configure_extraction_matrix(
            project_dir=project_dir,
            clusters=clusters,
            method_codes=method_codes,
            exclusion_codes=exclusion_codes,
            studies=studies,
            excluded_studies=excluded_studies,
            replace_existing_studies=replace_existing_studies,
        )
    )


@mcp.tool()
def export_blinded_coder_workbook(
    project_dir: str,
    co_coders: list[str] | None = None,
    overlap_ratio: float = 0.2,
    output_filename: str | None = None,
) -> str:
    """Stage 5a: Export a 100%-blinded human second-coder Excel workbook (zero Coder 1 leakage) with cluster-stratified shared overlap."""
    return _as_json(
        reliability.export_blinded_coder_workbook(
            project_dir=project_dir,
            co_coders=co_coders,
            overlap_ratio=overlap_ratio,
            output_filename=output_filename,
        )
    )


@mcp.tool()
def compute_interrater_reliability(
    project_dir: str,
    coder2_workbook_path: str | None = None,
    columns: list[str] | None = None,
) -> str:
    """Stage 5b: Compute % Agreement, Cohen's Kappa (κ), PABAK, Landis–Koch interpretation, and itemized disagreements between Coder 1 & Coder 2."""
    return _as_json(
        reliability.compute_interrater_reliability(
            project_dir=project_dir,
            coder2_workbook_path=coder2_workbook_path,
            columns=columns,
        )
    )


@mcp.tool()
def verify_and_export_bibliography(
    project_dir: str,
    verify_online: bool = False,
    extra_references: list[dict[str, str]] | None = None,
    timeout_seconds: float = 10.0,
) -> str:
    """Stage 6: Optionally verify DOIs via Crossref/OpenAlex and export clean Mendeley/Zotero .bib and .ris files with properly split AU lines."""
    return _as_json(
        bibliography.verify_and_export_bibliography(
            project_dir=project_dir,
            verify_online=verify_online,
            extra_references=extra_references,
            timeout_seconds=timeout_seconds,
        )
    )


@mcp.tool()
def generate_slr_figures(project_dir: str) -> str:
    """Stage 7a: Generate 3 publication-ready 300 DPI PNG figures directly from Search_Log.xlsx and Extraction_Matrix.xlsx."""
    return _as_json(sync_report.generate_slr_figures(project_dir=project_dir))


@mcp.tool()
def sync_manuscript_snapshot(
    project_dir: str,
    custom_abstract: str | None = None,
    regenerate_figures: bool = True,
    sync_protocol: bool = True,
) -> str:
    """Stage 7b: Build Single-Source-of-Truth synchronized DOCX & Markdown manuscript snapshot directly from Excel with embedded 300 DPI figures."""
    return _as_json(
        sync_report.sync_manuscript_snapshot(
            project_dir=project_dir,
            custom_abstract=custom_abstract,
            regenerate_figures=regenerate_figures,
            sync_protocol=sync_protocol,
        )
    )


@mcp.tool()
def run_four_way_integrity_audit(
    project_dir: str,
    save_report_json: bool = True,
) -> str:
    """Stage 8a: Run 20+ deterministic 4-way integrity checks across PDFs, Search Log, Extraction Matrix/Notes, Manuscript, Bib/RIS & Protocol."""
    return _as_json(
        audit.run_four_way_integrity_audit(
            project_dir=project_dir,
            save_report_json=save_report_json,
        )
    )


@mcp.tool()
def get_project_status(project_dir: str) -> str:
    """Stage 8b: Inspect current project completion metrics, pending queues, last audit verdict, and recommended next tool."""
    return _as_json(audit.get_project_status(project_dir=project_dir))


# =============================================================================
# 3 MCP RESOURCES
# =============================================================================


@mcp.resource("scholarflow://guide/spar4slr")
def resource_spar4slr_guide() -> str:
    """SPAR-4-SLR (Paul et al., 2021) & ADO-TCM coding framework quick reference."""
    return (
        "# SPAR-4-SLR & ADO-TCM Methodology Reference\n\n"
        "## 1. SPAR-4-SLR Three-Stage Architecture\n"
        "- **Assembling (Identification & Acquisition):** Combine 5 explicit pathways — "
        "(P1) Primary DB (e.g., Scopus), (P2) Secondary DBs (OpenAlex, Crossref, SciELO, GS/PoP), "
        "(P3) Complementary Web/Grey Literature, (P4) Initial Team Seed Corpus, and "
        "(P5) Backward & Forward Snowballing. Deduplicate deterministically.\n"
        "- **Arranging (Organization & Purification):** Screen titles/abstracts against pre-registered "
        "eligibility rules, retrieve OA PDFs legally, and log unretrieved items in a clickable manual queue.\n"
        "- **Assessing (Evaluation & Reporting):** Route verified full-text PDFs non-destructively, extract "
        "into closed-taxonomy columns + ADO-TCM, run blinded human co-coding (Cohen's Kappa & PABAK), "
        "synchronize manuscript/figures from Excel, and pass the 4-way deterministic integrity audit.\n\n"
        "## 2. ADO-TCM Synthesis Matrix\n"
        "- **ADO (What is studied):** Antecedents (A), Decisions (D), Outcomes (O).\n"
        "- **TCM (How it is studied):** Theories (T), Contexts (C), Methods (M).\n"
    )


@mcp.resource("scholarflow://guide/integrity-contract")
def resource_integrity_contract() -> str:
    """4-Way Deterministic Integrity Audit Contract."""
    return (
        "# 4-Way Deterministic Integrity Audit Contract\n\n"
        "Before submitting a systematic review or empirical manuscript, verify all 4 directions:\n"
        "1. **Physical Corpus:** Staging folder empty; 0 corrupt PDFs via PyMuPDF; 1-to-1 code ↔ filename match.\n"
        "2. **Search Log:** Exact 5-pathway Assembling, Arranging, and Assessing arithmetic; 100% clickable manual queue.\n"
        "3. **Extraction Matrix & Notes:** Unique codes; 100% closed-taxonomy completeness; Excel-safe cell-range DataValidation "
        "(0 inline >255-char formulas); 1-to-1 Markdown reading notes; exact Excluded_FullText count match.\n"
        "4. **Manuscript, Bib/RIS & Protocol:** Abstract word budget; embedded & prose-referenced 300 DPI figures; "
        "row-for-row appendix table sync; 0 forbidden legacy patterns; 100% prose citation coverage; "
        "clean .bib/.ris exports (0 '&'-joined AU lines); Protocol synchronized to locked corpus N.\n"
    )


@mcp.resource("scholarflow://templates/reading-note")
def resource_reading_note_template() -> str:
    """Standard 1-to-1 Markdown Reading Note Template."""
    return (
        "# [<Study_Code>] <Authors> (<Year>) — <Title>\n\n"
        "- **Study Code:** `<Study_Code>`\n"
        "- **DOI / Link:** <DOI>\n"
        "- **Paradigm Cluster:** <Paradigm_Cluster>\n"
        "- **Closed Method Code:** <Method_Code>\n\n"
        "## Methodology & Key Findings\n"
        "<Concise page-verified findings>\n\n"
        "## ADO–TCM Extraction\n"
        "- **Antecedents (A):** ...\n"
        "- **Decisions (D):** ...\n"
        "- **Outcomes (O):** ...\n"
        "- **Theories (T):** ...\n"
        "- **Contexts (C):** ...\n"
        "- **Methods (M):** ...\n"
    )


# =============================================================================
# 3 GUIDED MCP PROMPTS
# =============================================================================


@mcp.prompt()
def start_systematic_review(topic: str, language: str = "en", framework: str = "spar4slr") -> str:
    """Guided kickoff workflow: FINER RQ formulation, protocol confirmation, and workspace scaffolding."""
    return (
        f"You are orchestrating a reproducible systematic literature review on: '{topic}'.\n"
        f"Language preset: `{language}` | Protocol framework: `{framework}`.\n\n"
        "Follow these steps strictly:\n"
        "1. Formulate 2–3 FINER-scored research questions and propose explicit inclusion/exclusion rules "
        "(E1–E4) plus a closed coding taxonomy (Paradigm Clusters + Method Codes M1–Mn).\n"
        "2. Pause and ask the researcher to confirm the protocol before touching any records.\n"
        "3. Once confirmed, call `init_research_project` to scaffold the reproducible workspace.\n"
    )


@mcp.prompt()
def run_screening_and_extraction_cycle(project_dir: str) -> str:
    """Guided prompt for multi-DB ingestion, hybrid OA/manual PDF retrieval, routing, and blinded reliability check."""
    return (
        f"Continue the reproducible SLR cycle for project: `{project_dir}`.\n\n"
        "1. Call `get_project_status` first to inspect current stage counts.\n"
        "2. Ingest database exports with `ingest_and_deduplicate_exports` and verify flow arithmetic with `reconcile_flow_counts`.\n"
        "3. Retrieve legal OA PDFs via `fetch_open_access_pdfs` (never bypass CAPTCHAs) and log remaining items via `build_manual_download_queue`.\n"
        "4. Route reviewed full texts with `route_corpus_files` and lock closed codes via `configure_extraction_matrix`.\n"
        "5. Export a zero-leakage blinded workbook via `export_blinded_coder_workbook`, then compute κ & PABAK via `compute_interrater_reliability`.\n"
    )


@mcp.prompt()
def pre_submission_integrity_gate(project_dir: str) -> str:
    """Pre-submission gate: verify BibTeX/RIS, sync data-driven manuscript/figures, and run 4-way integrity audit."""
    return (
        f"Execute the mandatory pre-submission integrity gate for `{project_dir}`:\n\n"
        "1. Export and verify bibliography via `verify_and_export_bibliography`.\n"
        "2. Regenerate 300 DPI figures and sync the Single-Source-of-Truth manuscript via `sync_manuscript_snapshot`.\n"
        "3. Run `run_four_way_integrity_audit`.\n"
        "4. If ANY check fails, report exact failing IDs and remediate before proceeding to simulated 5-reviewer peer review.\n"
    )


def main() -> None:
    """CLI entrypoint for `scholarflow-mcp` (stdio transport)."""
    mcp.run()


if __name__ == "__main__":
    main()
