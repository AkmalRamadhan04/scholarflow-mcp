"""Tests for Stage 2 (OA HTML extraction, CAPTCHA refusal, clickable queue) & Stage 3 (corpus routing)."""

from __future__ import annotations

from pathlib import Path

import openpyxl

from scholarflow_mcp.core import acquisition, corpus, scaffold
from tests.conftest import create_valid_pdf


def test_html_pdf_extraction_and_captcha_refusal() -> None:
    ojs_html = (
        '<html><head><meta name="citation_pdf_url" '
        'content="https://journal.example.org/article/download/42/99"></head></html>'
    )
    pdf_url, blocked = acquisition.extract_pdf_url_from_html(ojs_html, "https://journal.example.org/article/view/42")
    assert blocked is False
    assert pdf_url == "https://journal.example.org/article/download/42/99"

    captcha_html = "<html><body>Please complete the cf-turnstile challenge to verify you are human.</body></html>"
    pdf_url2, blocked2 = acquisition.extract_pdf_url_from_html(captcha_html, "https://blocked.example.org")
    assert blocked2 is True
    assert pdf_url2 is None


def test_clickable_manual_download_queue(tmp_path: Path) -> None:
    proj = tmp_path / "slr"
    scaffold.init_research_project(str(proj), "Queue Test", language="en")

    res = acquisition.build_manual_download_queue(
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
                "url": "",
            }
        ],
    )
    assert res["queue_count"] == 1
    assert res["clickable_links_count"] == 1

    wb = openpyxl.load_workbook(proj / "03_Extraction_Coding" / "Search_Log.xlsx")
    ws = wb["Fulltext_Download_Queue"]
    assert ws.cell(2, 11).hyperlink is not None
    assert ws.cell(2, 11).hyperlink.target == "https://doi.org/10.1000/demo.2023.008"


def test_route_corpus_files_and_corrupt_pdf_guard(tmp_path: Path) -> None:
    proj = tmp_path / "slr"
    scaffold.init_research_project(str(proj), "Routing Test", language="en")
    unclass = proj / "02_Literature" / "2_Unclassified"

    create_valid_pdf(unclass / "raw_s01.pdf", "Valid Study S01")
    create_valid_pdf(unclass / "raw_excl.pdf", "Excluded Editorial")
    create_valid_pdf(unclass / "raw_mis.pdf", "Botany misfiltered")
    (unclass / "corrupt.pdf").write_bytes(b"NOT A REAL PDF FILE")

    res = corpus.route_corpus_files(
        str(proj),
        [
            {
                "source_filename": "raw_s01.pdf",
                "destination": "core",
                "study_code": "S01",
                "authors": "Alvarez, M.; Chen, L.",
                "year": "2021",
                "title": "Spatial Hedonic Pricing of Urban Pocket Parks",
                "doi": "10.1000/demo.2021.001",
                "paradigm_cluster": "Cluster A: Market & Statutory Appraisal",
                "method_code": "M1: Market Comparison / Hedonic",
            },
            {
                "source_filename": "raw_excl.pdf",
                "destination": "excluded",
                "study_code": "EX01",
                "authors": "Ericson, H.",
                "year": "2020",
                "title": "Editorial Note on Urban Aesthetics",
                "exclusion_code": "E3",
                "reason": "Non-primary editorial note",
            },
            {
                "source_filename": "raw_mis.pdf",
                "destination": "misfiltered",
                "reason": "Botanical species inventory outside valuation domain",
            },
            {
                "source_filename": "corrupt.pdf",
                "destination": "core",
                "study_code": "S99",
                "authors": "Broken, X.",
                "year": "2025",
                "title": "Corrupt File",
            },
        ],
    )
    assert res["status"] == "partial_with_corrupt_pdfs"
    assert res["corrupt_pdfs_count"] == 1
    assert res["routed_count"] == 3
    assert res["notes_created_count"] == 1
    assert (proj / "02_Literature" / "1_Core" / "S01_Alvarez_2021.pdf").exists()
    assert not (proj / "02_Literature" / "1_Core" / "S99_Broken_2025.pdf").exists()
    assert (proj / "03_Extraction_Coding" / "Reading_Notes" / "S01_Alvarez_2021.md").exists()
    assert "EX01" in (proj / "03_Extraction_Coding" / "Reading_Notes" / "_Excluded_FullText.md").read_text(encoding="utf-8")
    assert "raw_mis.pdf" in (proj / "02_Literature" / "archive_misfiltered" / "_Manifest_Misfiltered.csv").read_text(encoding="utf-8")
