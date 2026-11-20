"""Tests for Stage 0 (project scaffolding) and Stage 1 (multi-DB dedup & SPAR-4-SLR flow)."""

from __future__ import annotations

from pathlib import Path

from scholarflow_mcp.core import records, scaffold


def test_init_project_en_and_id_presets(tmp_path: Path) -> None:
    en_dir = tmp_path / "proj_en"
    res_en = scaffold.init_research_project(
        project_dir=str(en_dir),
        project_title="Synthetic Urban Green Valuation SLR",
        authors=["Akmal Ramadhan"],
        language="en",
        framework="spar4slr",
    )
    assert res_en["status"] == "ok"
    assert (en_dir / "scholarflow.yaml").exists()
    assert (en_dir / "02_Literature" / "1_Core").is_dir()
    assert (en_dir / "03_Extraction_Coding" / "Search_Log.xlsx").exists()
    assert (en_dir / "03_Extraction_Coding" / "Extraction_Matrix.xlsx").exists()
    assert (en_dir / "00_Admin" / "Review_Protocol.docx").exists()

    id_dir = tmp_path / "proj_id"
    res_id = scaffold.init_research_project(
        project_dir=str(id_dir),
        project_title="Tinjauan Sistematis Valuasi Ruang Hijau",
        authors=["Akmal Ramadhan"],
        language="id",
        framework="spar4slr",
    )
    assert res_id["status"] == "ok"
    assert (id_dir / "02_Literatur" / "1_Inti").is_dir()
    assert (id_dir / "03_Ekstraksi_Koding" / "Log_Pencarian.xlsx").exists()
    assert (id_dir / "03_Ekstraksi_Koding" / "Matriks_Induk.xlsx").exists()
    assert (id_dir / "00_Admin" / "Protokol_Tinjauan.docx").exists()


def test_ingest_dedup_csv_ris_bib_and_reconcile(tmp_path: Path) -> None:
    proj = tmp_path / "slr"
    scaffold.init_research_project(str(proj), "Demo SLR", language="en")

    repo_root = Path(__file__).resolve().parents[1]
    scopus_csv = repo_root / "examples" / "synthetic_demo" / "sample_exports" / "scopus_sample.csv"
    openalex_csv = repo_root / "examples" / "synthetic_demo" / "sample_exports" / "openalex_sample.csv"

    # Also test RIS and BibTeX ingestion with 1 duplicate DOI and 1 unique snowball record
    ris_file = tmp_path / "snowball.ris"
    ris_file.write_text(
        "TY  - JOUR\n"
        "AU  - Alvarez, M.\n"
        "AU  - Chen, L.\n"
        "PY  - 2021\n"
        "TI  - Spatial Hedonic Pricing of Urban Pocket Parks in Mid-Sized Cities\n"
        "DO  - 10.1000/demo.2021.001\n"
        "ER  - \n\n"
        "TY  - JOUR\n"
        "AU  - Tanaka, K.\n"
        "PY  - 2025\n"
        "TI  - Snowball Study on Municipal Greenway Assessment\n"
        "DO  - 10.1000/demo.2025.010\n"
        "ER  - \n",
        encoding="utf-8",
    )

    res = records.ingest_and_deduplicate_exports(
        str(proj),
        [
            {"path": str(scopus_csv), "pathway": "pathway_1_primary_db", "source_db": "Scopus"},
            {"path": str(openalex_csv), "pathway": "pathway_2_secondary_db", "source_db": "OpenAlex"},
            {"path": str(ris_file), "pathway": "pathway_5_snowballing", "source_db": "Snowballing"},
        ],
    )
    # 5 + 5 + 2 = 12 raw; 2 duplicates of demo.2021.001 removed -> 10 unique
    assert res["raw_records_total"] == 12
    assert res["duplicates_removed"] == 2
    assert res["unique_records_total"] == 10
    assert res["pathway_unique_counts"]["pathway_1_primary_db"] == 5
    assert res["pathway_unique_counts"]["pathway_2_secondary_db"] == 4
    assert res["pathway_unique_counts"]["pathway_5_snowballing"] == 1

    # Valid reconciliation: 10 - 2 TA excl = 8 cand; 8 - 1 unretrieved = 7 assessed; 7 - 1 FT excl = 6 core
    rec_ok = records.reconcile_flow_counts(
        str(proj),
        {
            "excluded_title_abstract": 2,
            "candidates_passing_ta": 8,
            "unretrieved_fulltext": 1,
            "fulltext_assessed": 7,
            "excluded_fulltext": 1,
            "final_core_corpus": 6,
        },
    )
    assert rec_ok["status"] == "PASS"

    # Broken arithmetic must fail
    rec_bad = records.reconcile_flow_counts(str(proj), {"final_core_corpus": 5})
    assert rec_bad["status"] == "FAIL"
