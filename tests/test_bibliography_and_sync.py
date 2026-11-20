"""Tests for Stage 6 (BibTeX/RIS clean export) and Stage 7 (300 DPI figures + DOCX/MD manuscript sync)."""

from __future__ import annotations

from pathlib import Path

from docx import Document

from scholarflow_mcp.core import bibliography, coding, records, scaffold, sync_report


def test_bibliography_author_splitting_and_sync_manuscript(
    tmp_path: Path, synthetic_studies: list[dict[str, str]]
) -> None:
    proj = tmp_path / "slr"
    scaffold.init_research_project(str(proj), "Sync & Bib Test", language="en")
    coding.configure_extraction_matrix(str(proj), studies=synthetic_studies, replace_existing_studies=True)
    records.reconcile_flow_counts(
        str(proj),
        {
            "pathway_1_primary_db": 5,
            "pathway_2_secondary_db": 4,
            "pathway_5_snowballing": 1,
            "total_unique_screened": 10,
            "excluded_title_abstract": 2,
            "candidates_passing_ta": 8,
            "unretrieved_fulltext": 1,
            "fulltext_assessed": 7,
            "excluded_fulltext": 1,
            "final_core_corpus": 6,
        },
    )

    # Test author splitter against '&' bugs
    assert bibliography.split_authors_clean("Smith, A. & Jones, B.") == ["Smith, A.", "Jones, B."]

    bib_res = bibliography.verify_and_export_bibliography(str(proj), verify_online=False)
    assert bib_res["entries_exported"] == 6
    ris_text = Path(bib_res["ris_path"]).read_text(encoding="utf-8")
    au_lines = [ln for ln in ris_text.splitlines() if ln.startswith("AU  -")]
    assert len(au_lines) >= 6
    assert all("&" not in ln for ln in au_lines)

    sync_res = sync_report.sync_manuscript_snapshot(str(proj))
    assert sync_res["locked_corpus_n"] == 6
    assert sync_res["figures_embedded"] == 3
    assert 150 <= sync_res["abstract_words"] <= 300
    assert Path(sync_res["manuscript_docx"]).exists()
    assert Path(sync_res["manuscript_md"]).exists()

    doc = Document(sync_res["manuscript_docx"])
    assert len(doc.inline_shapes) == 3
    assert len(doc.tables[-1].rows) == 7  # header + 6 studies
