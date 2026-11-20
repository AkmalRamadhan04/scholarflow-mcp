"""End-to-end 8-stage synthetic pipeline test, negative integrity regressions & FastMCP server check."""

from __future__ import annotations

import asyncio
import json
from pathlib import Path

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
from scholarflow_mcp.server import mcp
from tests.conftest import create_valid_pdf


def test_full_8_stage_e2e_and_four_way_audit_pass_and_fail(
    tmp_path: Path, synthetic_studies: list[dict[str, str]]
) -> None:
    proj = tmp_path / "full_e2e"
    repo_root = Path(__file__).resolve().parents[1]
    scopus_csv = repo_root / "examples" / "synthetic_demo" / "sample_exports" / "scopus_sample.csv"
    openalex_csv = repo_root / "examples" / "synthetic_demo" / "sample_exports" / "openalex_sample.csv"

    # Stage 0: Scaffold
    scaffold.init_research_project(
        str(proj),
        project_title="Urban Green Infrastructure Valuation: A Reproducible SPAR-4-SLR",
        authors=["Akmal Ramadhan"],
        language="en",
        framework="spar4slr",
    )

    # Stage 1: Multi-DB Ingest & Reconcile (5 + 4 = 9 unique; 1 TA excl -> 8; 1 unretrieved -> 7; 1 FT excl -> 6 core)
    ing = records.ingest_and_deduplicate_exports(
        str(proj),
        [
            {"path": str(scopus_csv), "pathway": "pathway_1_primary_db", "source_db": "Scopus"},
            {"path": str(openalex_csv), "pathway": "pathway_2_secondary_db", "source_db": "OpenAlex"},
        ],
    )
    assert ing["unique_records_total"] == 9
    rec = records.reconcile_flow_counts(
        str(proj),
        {
            "excluded_title_abstract": 1,
            "candidates_passing_ta": 8,
            "unretrieved_fulltext": 1,
            "fulltext_assessed": 7,
            "excluded_fulltext": 1,
            "final_core_corpus": 6,
        },
    )
    assert rec["status"] == "PASS"

    # Stage 2: Manual queue for 1 unretrieved study
    q_res = acquisition.build_manual_download_queue(
        str(proj),
        queue_items=[
            {
                "record_id": "R0008",
                "pathway": "pathway_2_secondary_db",
                "priority": "High",
                "authors": "Haddad, L.",
                "year": "2023",
                "title": "Restricted Working Paper on Municipal Park Bonds",
                "venue": "Archive Working Series",
                "doi": "10.1000/demo.2023.008",
            }
        ],
    )
    assert q_res["clickable_links_count"] == 1

    # Stage 3: Stage 6 core PDFs + 1 excluded PDF + 1 misfiltered PDF, then route non-destructively
    unclass = proj / "02_Literature" / "2_Unclassified"
    decisions = []
    for st in synthetic_studies:
        raw_name = f"raw_{st['study_code']}.pdf"
        create_valid_pdf(unclass / raw_name, f"{st['study_code']} - {st['title']}")
        decisions.append(
            {
                "source_filename": raw_name,
                "destination": "core",
                **st,
            }
        )
    create_valid_pdf(unclass / "raw_editorial.pdf", "Editorial excluded")
    decisions.append(
        {
            "source_filename": "raw_editorial.pdf",
            "destination": "excluded",
            "study_code": "EX01",
            "authors": "Ericson, H.",
            "year": "2020",
            "title": "Editorial Note on Urban Aesthetics",
            "exclusion_code": "E3",
            "reason": "Editorial commentary",
        }
    )
    route_res = corpus.route_corpus_files(str(proj), decisions)
    assert route_res["corrupt_pdfs_count"] == 0
    assert route_res["notes_created_count"] == 6

    # Attach exact reading_note_path & pdf_filename returned by routing
    routed_map = {r["study_code"]: r for r in route_res["routed"] if r["destination"] == "core"}
    enriched_studies = []
    for st in synthetic_studies:
        r_info = routed_map[st["study_code"]]
        enriched_studies.append(
            {
                **st,
                "reading_note_path": r_info["reading_note_rel"],
                "pdf_filename": Path(r_info["target_path"]).name,
            }
        )

    # Stage 4: Extraction Matrix + Excluded_FullText
    coding.configure_extraction_matrix(
        str(proj),
        studies=enriched_studies,
        excluded_studies=[
            {
                "record_id": "EX01",
                "authors": "Ericson, H.",
                "year": "2020",
                "title": "Editorial Note on Urban Aesthetics",
                "doi": "10.1000/demo.2020.005",
                "exclusion_code": "E3",
                "reason_detail": "Non-primary editorial",
                "category": "Excluded",
            }
        ],
        replace_existing_studies=True,
    )

    # Stage 5: Blinded workbook export
    blind = reliability.export_blinded_coder_workbook(str(proj), co_coders=["Reviewer_1", "Reviewer_2"])
    assert blind["coder1_labels_leaked"] is False

    # Stage 6: Clean Bibliography export
    bibliography.verify_and_export_bibliography(str(proj), verify_online=False)

    # Stage 7: 300 DPI figures + Single-Source-of-Truth Manuscript & Protocol sync
    sync_report.sync_manuscript_snapshot(str(proj))

    # Stage 8: 4-Way Deterministic Integrity Audit MUST PASS 100%!
    report = audit.run_four_way_integrity_audit(str(proj))
    assert report["status"] == "PASS", f"Unexpected failures: {report['failed_checks']}"
    assert report["total_checks"] == 25
    assert report["passed"] == 25
    assert report["failed"] == 0

    # Status dashboard confirms completion
    status = audit.get_project_status(str(proj))
    assert status["last_audit"] == "PASS (25/25)"

    # Negative regression 1: Corrupt PDF in 1_Core must trigger D1_03 FAIL
    (proj / "02_Literature" / "1_Core" / "S99_Corrupt.pdf").write_bytes(b"corrupt")
    report_corrupt = audit.run_four_way_integrity_audit(str(proj))
    assert report_corrupt["status"] == "FAIL"
    failed_ids = {c["id"] for c in report_corrupt["failed_checks"]}
    assert "D1_03_all_pdfs_readable" in failed_ids
    (proj / "02_Literature" / "1_Core" / "S99_Corrupt.pdf").unlink()

    # Negative regression 2: Malformed '&' inside RIS AU line must trigger D4_08 FAIL
    ris_path = proj / "04_Tools" / "references.ris"
    orig_ris = ris_path.read_text(encoding="utf-8")
    ris_path.write_text(orig_ris.replace("AU  - Alvarez, M.", "AU  - Alvarez, M. & Chen, L."), encoding="utf-8")
    report_ris = audit.run_four_way_integrity_audit(str(proj))
    assert report_ris["status"] == "FAIL"
    assert "D4_08_bib_and_ris_clean" in {c["id"] for c in report_ris["failed_checks"]}


def test_fastmcp_server_registration_and_tool_call(tmp_path: Path) -> None:
    tools = asyncio.run(mcp.list_tools())
    resources = asyncio.run(mcp.list_resources())
    prompts = asyncio.run(mcp.list_prompts())

    assert len(tools) == 14
    assert len(resources) == 3
    assert len(prompts) == 3

    tool_names = {t.name for t in tools}
    assert "init_research_project" in tool_names
    assert "export_blinded_coder_workbook" in tool_names
    assert "compute_interrater_reliability" in tool_names
    assert "run_four_way_integrity_audit" in tool_names

    # Invoke a real MCP tool through MCPServer (supports mcp v2 CallToolResult and v1 list/tuple)
    mcp_proj = tmp_path / "mcp_call_test"
    result = asyncio.run(
        mcp.call_tool(
            "init_research_project",
            {"project_dir": str(mcp_proj), "project_title": "MCP Tool Smoke Test"},
        )
    )
    if hasattr(result, "content"):
        blocks = result.content
    elif isinstance(result, tuple):
        blocks = result[0]
    else:
        blocks = result
    payload = json.loads(blocks[0].text)
    assert payload["status"] == "ok"
    assert (mcp_proj / "scholarflow.yaml").exists()
