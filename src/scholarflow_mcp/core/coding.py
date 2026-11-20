"""Stage 4: Closed-taxonomy Extraction Matrix, ADO-TCM columns & Excel-safe DataValidation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl

from scholarflow_mcp.config import ProjectConfig
from scholarflow_mcp.core.excel_utils import add_safe_range_dropdown, style_sheet


def configure_extraction_matrix(
    project_dir: str,
    clusters: list[str] | None = None,
    method_codes: list[str] | None = None,
    exclusion_codes: list[str] | None = None,
    studies: list[dict[str, Any]] | None = None,
    excluded_studies: list[dict[str, Any]] | None = None,
    replace_existing_studies: bool = False,
) -> dict[str, Any]:
    """Configure closed taxonomy codebooks, safe Excel dropdowns, and populate Extraction_Matrix.xlsx."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()
    matrix_path = cfg.resolve_path(root, "extraction_matrix_file")

    updated_cfg = False
    if clusters is not None:
        cfg.clusters = list(clusters)
        updated_cfg = True
    if method_codes is not None:
        cfg.method_codes = list(method_codes)
        updated_cfg = True
    if exclusion_codes is not None:
        cfg.exclusion_codes = list(exclusion_codes)
        updated_cfg = True
    if updated_cfg:
        cfg.save(root)

    wb = openpyxl.load_workbook(matrix_path)

    # 1. Refresh Codebook sheet
    ws_cb = wb[preset.sheet_codebook]
    if ws_cb.max_row > 1:
        ws_cb.delete_rows(2, ws_cb.max_row - 1)
    max_len = max(len(cfg.clusters), len(cfg.method_codes), len(cfg.exclusion_codes), 1)
    for i in range(max_len):
        ws_cb.append(
            [
                cfg.clusters[i] if i < len(cfg.clusters) else "",
                cfg.method_codes[i] if i < len(cfg.method_codes) else "",
                cfg.exclusion_codes[i] if i < len(cfg.exclusion_codes) else "",
            ]
        )
    style_sheet(ws_cb, {1: 42, 2: 42, 3: 44})

    # 2. Upsert studies in Extraction_Matrix
    ws_m = wb[preset.sheet_matrix]
    headers = [c.value for c in ws_m[1]]
    col_map = {h: idx + 1 for idx, h in enumerate(headers) if h}

    if replace_existing_studies and ws_m.max_row > 1:
        ws_m.delete_rows(2, ws_m.max_row - 1)

    existing_rows: dict[str, int] = {}
    for r in range(2, ws_m.max_row + 1):
        code_val = ws_m.cell(r, 1).value
        if code_val:
            existing_rows[str(code_val).strip()] = r

    valid_clusters = set(cfg.clusters)
    valid_methods = set(cfg.method_codes)
    validation_warnings: list[str] = []

    if studies:
        for st in studies:
            code = str(st["study_code"]).strip()
            cl = str(st.get("paradigm_cluster", "")).strip()
            mc = str(st.get("method_code", "")).strip()
            if cl and valid_clusters and cl not in valid_clusters:
                validation_warnings.append(f"{code}: paradigm_cluster '{cl}' not in closed codebook")
            if mc and valid_methods and mc not in valid_methods:
                validation_warnings.append(f"{code}: method_code '{mc}' not in closed codebook")

            field_values = {
                "Study_Code": code,
                "Authors": st.get("authors", ""),
                "Year": st.get("year", ""),
                "Title": st.get("title", ""),
                "Country_Context": st.get("country_context", ""),
                "Venue": st.get("venue", ""),
                "DOI": st.get("doi", ""),
                "Paradigm_Cluster": cl,
                "Method_Code": mc,
                "Sample_or_Object": st.get("sample_or_object", ""),
                "Key_Findings": st.get("key_findings", ""),
                "Reading_Note_Path": st.get("reading_note_path", ""),
                "PDF_Filename": st.get("pdf_filename", ""),
                "ADO_Antecedents": st.get("ado_antecedents", ""),
                "ADO_Decisions": st.get("ado_decisions", ""),
                "ADO_Outcomes": st.get("ado_outcomes", ""),
                "TCM_Theories": st.get("tcm_theories", ""),
                "TCM_Contexts": st.get("tcm_contexts", ""),
                "TCM_Methods_Detail": st.get("tcm_methods_detail", ""),
            }
            if code in existing_rows:
                r_idx = existing_rows[code]
                for h_name, val in field_values.items():
                    if h_name in col_map:
                        ws_m.cell(r_idx, col_map[h_name]).value = val
            else:
                row_out = [field_values.get(h, "") for h in headers]
                ws_m.append(row_out)
                existing_rows[code] = ws_m.max_row

    # Refresh safe cell-range DataValidation on Extraction_Matrix & Coder2_Worksheet
    ws_m.data_validations.dataValidation.clear()
    end_row_m = max(ws_m.max_row + 50, 200)
    if cfg.clusters and "Paradigm_Cluster" in col_map:
        cl_letter = openpyxl.utils.get_column_letter(col_map["Paradigm_Cluster"])
        add_safe_range_dropdown(
            ws_m, cl_letter, 2, end_row_m, preset.sheet_codebook, "A", 2, 1 + len(cfg.clusters),
            prompt_title="Paradigm Cluster",
        )
    if cfg.method_codes and "Method_Code" in col_map:
        mc_letter = openpyxl.utils.get_column_letter(col_map["Method_Code"])
        add_safe_range_dropdown(
            ws_m, mc_letter, 2, end_row_m, preset.sheet_codebook, "B", 2, 1 + len(cfg.method_codes),
            prompt_title="Closed Method Code",
        )
    style_sheet(ws_m)

    # 3. Sync Coder2_Worksheet rows (metadata only; preserves existing Coder 2 entries if present)
    ws_c2 = wb[preset.sheet_coder2]
    c2_existing: dict[str, tuple[Any, ...]] = {}
    for row in list(ws_c2.iter_rows(values_only=True))[1:]:
        if row and row[0]:
            c2_existing[str(row[0]).strip()] = row
    if ws_c2.max_row > 1:
        ws_c2.delete_rows(2, ws_c2.max_row - 1)
    for r in range(2, ws_m.max_row + 1):
        code = str(ws_m.cell(r, 1).value or "").strip()
        if not code:
            continue
        prev = c2_existing.get(code)
        ws_c2.append(
            [
                code,
                ws_m.cell(r, col_map["Authors"]).value,
                ws_m.cell(r, col_map["Year"]).value,
                ws_m.cell(r, col_map["Title"]).value,
                prev[4] if prev and len(prev) > 4 else "",
                prev[5] if prev and len(prev) > 5 else "",
                prev[6] if prev and len(prev) > 6 else "",
                prev[7] if prev and len(prev) > 7 else "Pending",
                prev[8] if prev and len(prev) > 8 else "",
            ]
        )
    ws_c2.data_validations.dataValidation.clear()
    end_row_c2 = max(ws_c2.max_row, 2)
    if cfg.clusters:
        add_safe_range_dropdown(
            ws_c2, "F", 2, end_row_c2, preset.sheet_codebook, "A", 2, 1 + len(cfg.clusters),
            prompt_title="Paradigm Cluster",
        )
    if cfg.method_codes:
        add_safe_range_dropdown(
            ws_c2, "G", 2, end_row_c2, preset.sheet_codebook, "B", 2, 1 + len(cfg.method_codes),
            prompt_title="Closed Method Code",
        )
    style_sheet(ws_c2, {1: 14, 2: 26, 3: 8, 4: 44, 5: 20, 6: 36, 7: 34, 8: 20, 9: 32})

    # 4. Populate Excluded Full-Text sheet if supplied
    ws_ex = wb[preset.sheet_excluded_ft]
    if excluded_studies is not None:
        if replace_existing_studies and ws_ex.max_row > 1:
            ws_ex.delete_rows(2, ws_ex.max_row - 1)
        for ex in excluded_studies:
            ws_ex.append(
                [
                    ex.get("record_id", ""),
                    ex.get("authors", ""),
                    ex.get("year", ""),
                    ex.get("title", ""),
                    ex.get("doi", ""),
                    ex.get("exclusion_code", ""),
                    ex.get("reason_detail", ""),
                    ex.get("category", "Excluded"),
                ]
            )
        style_sheet(ws_ex)

    wb.save(matrix_path)

    total_studies = sum(1 for r in range(2, ws_m.max_row + 1) if ws_m.cell(r, 1).value)
    total_excluded = sum(1 for r in range(2, ws_ex.max_row + 1) if ws_ex.cell(r, 1).value)

    return {
        "status": "ok",
        "matrix_path": str(matrix_path),
        "total_core_studies": total_studies,
        "total_excluded_fulltext": total_excluded,
        "clusters_count": len(cfg.clusters),
        "method_codes_count": len(cfg.method_codes),
        "validation_warnings": validation_warnings,
    }
