"""Tests for Stage 4 (closed taxonomy + safe Excel DataValidation) & Stage 5 (blinded export + Kappa/PABAK)."""

from __future__ import annotations

from pathlib import Path

import openpyxl

from scholarflow_mcp.core import coding, excel_utils, reliability, scaffold


def test_closed_taxonomy_safe_dropdowns_and_blinded_reliability(
    tmp_path: Path, synthetic_studies: list[dict[str, str]]
) -> None:
    proj = tmp_path / "slr"
    scaffold.init_research_project(str(proj), "Reliability Test", language="en")

    res_mat = coding.configure_extraction_matrix(
        str(proj),
        studies=synthetic_studies,
        replace_existing_studies=True,
    )
    assert res_mat["total_core_studies"] == 6
    assert res_mat["validation_warnings"] == []

    # Verify zero >255-char inline DataValidation formulas
    wb_m = openpyxl.load_workbook(proj / "03_Extraction_Coding" / "Extraction_Matrix.xlsx")
    assert excel_utils.detect_unsafe_inline_validations(wb_m["Extraction_Matrix"]) == []
    assert excel_utils.detect_unsafe_inline_validations(wb_m["Coder2_Worksheet"]) == []

    # Export 100%-blinded workbook for 2 co-coders with 33% overlap
    res_blind = reliability.export_blinded_coder_workbook(
        str(proj),
        co_coders=["CoCoder_A", "CoCoder_B"],
        overlap_ratio=0.33,
    )
    assert res_blind["coder1_labels_leaked"] is False
    assert res_blind["total_studies"] == 6
    assert res_blind["shared_overlap_count"] == 2  # 1 from Cluster A + 1 from Cluster B

    blind_path = Path(res_blind["blinded_workbook_path"])
    wb_b = openpyxl.load_workbook(blind_path)
    ws_b = wb_b["Blinded_Coding_Sheet"]
    # Confirm Paradigm_Cluster (col 8) and Method_Code (col 9) are completely blank (no leakage!)
    for r in range(2, ws_b.max_row + 1):
        assert ws_b.cell(r, 8).value in (None, "")
        assert ws_b.cell(r, 9).value in (None, "")

    # Simulate Coder 2 filling ratings: 6/6 exact match on Paradigm_Cluster, 5/6 on Method_Code (1 disagreement)
    for r, st in enumerate(synthetic_studies, start=2):
        ws_b.cell(r, 8).value = st["paradigm_cluster"]
        if st["study_code"] == "S05":
            ws_b.cell(r, 9).value = "M1: Market Comparison / Hedonic"  # intentional disagreement vs M7
        else:
            ws_b.cell(r, 9).value = st["method_code"]
    wb_b.save(blind_path)

    rel = reliability.compute_interrater_reliability(str(proj), str(blind_path))
    assert rel["unrated_by_coder2_count"] == 0
    assert rel["metrics"]["Paradigm_Cluster"]["percent_agreement"] == 100.0
    assert rel["metrics"]["Paradigm_Cluster"]["cohens_kappa"] == 1.0
    assert rel["metrics"]["Paradigm_Cluster"]["pabak"] == 1.0
    assert rel["metrics"]["Method_Code"]["percent_agreement"] == 83.33
    assert rel["metrics"]["Method_Code"]["cohens_kappa"] > 0.75
    assert rel["disagreements_count"] == 1
    assert rel["disagreements"][0]["study_code"] == "S05"
