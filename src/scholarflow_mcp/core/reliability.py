"""Stage 5: 100%-Blinded Second-Coder Excel Exporter & Inter-Rater Reliability (Cohen's Kappa & PABAK)."""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any

import openpyxl

from scholarflow_mcp.config import ProjectConfig
from scholarflow_mcp.core.excel_utils import (
    OVERLAP_FILL,
    add_safe_range_dropdown,
    style_sheet,
)


def cohens_kappa_and_pabak(
    y1: list[str],
    y2: list[str],
) -> dict[str, Any]:
    """Compute exact % Agreement, Cohen's Kappa (κ), PABAK, and Landis–Koch interpretation."""
    if len(y1) != len(y2):
        raise ValueError(f"Coder label vectors have mismatched lengths: {len(y1)} vs {len(y2)}")
    n = len(y1)
    if n == 0:
        return {
            "n": 0,
            "agreements": 0,
            "percent_agreement": None,
            "expected_agreement": None,
            "cohens_kappa": None,
            "pabak": None,
            "interpretation": "No paired ratings",
        }

    agreements = sum(1 for a, b in zip(y1, y2, strict=True) if a == b)
    po = agreements / n
    labels = sorted(set(y1) | set(y2))
    c1 = Counter(y1)
    c2 = Counter(y2)
    pe = sum((c1[lbl] / n) * (c2[lbl] / n) for lbl in labels)

    if abs(1.0 - pe) < 1e-12:
        kappa = 1.0 if abs(po - 1.0) < 1e-12 else 0.0
    else:
        kappa = (po - pe) / (1.0 - pe)

    # Prevalence-Adjusted Bias-Adjusted Kappa (Byrt et al., 1993) for k >= 2 categories
    k = max(len(labels), 2)
    pabak = (k * po - 1.0) / (k - 1.0) if k > 2 else (2.0 * po - 1.0)

    return {
        "n": n,
        "agreements": agreements,
        "percent_agreement": round(po * 100.0, 2),
        "expected_agreement": round(pe * 100.0, 2),
        "cohens_kappa": round(kappa, 4),
        "pabak": round(pabak, 4),
        "interpretation": interpret_landis_koch(kappa),
    }


def interpret_landis_koch(kappa: float | None) -> str:
    if kappa is None:
        return "N/A"
    if kappa < 0.0:
        return "Poor (< 0.00)"
    if kappa <= 0.20:
        return "Slight (0.00–0.20)"
    if kappa <= 0.40:
        return "Fair (0.21–0.40)"
    if kappa <= 0.60:
        return "Moderate (0.41–0.60)"
    if kappa <= 0.80:
        return "Substantial (0.61–0.80)"
    return "Almost Perfect (0.81–1.00)"


def export_blinded_coder_workbook(
    project_dir: str,
    co_coders: list[str] | None = None,
    overlap_ratio: float = 0.2,
    output_filename: str | None = None,
) -> dict[str, Any]:
    """Export a 100%-blinded standalone Excel workbook for independent human co-coders.

    Key guarantees:
    - Strips ALL Coder 1 labels (zero leakage of Paradigm_Cluster or Method_Code).
    - Stratifies shared overlap across paradigm clusters silently when multiple co-coders are listed.
    - Uses Excel-safe cell-range DataValidation referencing the bundled Codebook sheet.
    """
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()
    matrix_path = cfg.resolve_path(root, "extraction_matrix_file")
    out_path = (
        (cfg.resolve_path(root, "extraction_dir") / output_filename)
        if output_filename
        else cfg.resolve_path(root, "blinded_coder_file")
    )

    wb_src = openpyxl.load_workbook(matrix_path, data_only=True)
    ws_m = wb_src[preset.sheet_matrix]
    ws_cb_src = wb_src[preset.sheet_codebook]

    headers = [c.value for c in ws_m[1]]
    col = {h: i for i, h in enumerate(headers) if h}

    studies: list[dict[str, Any]] = []
    by_cluster: dict[str, list[str]] = {}
    for row in list(ws_m.iter_rows(values_only=True))[1:]:
        if not row or not row[0]:
            continue
        code = str(row[col["Study_Code"]]).strip()
        cl = str(row[col.get("Paradigm_Cluster", 7)] or "Unassigned").strip()
        studies.append(
            {
                "code": code,
                "authors": row[col.get("Authors", 1)] or "",
                "year": row[col.get("Year", 2)] or "",
                "title": row[col.get("Title", 3)] or "",
                "venue": row[col.get("Venue", 5)] or "",
                "doi": row[col.get("DOI", 6)] or "",
            }
        )
        by_cluster.setdefault(cl, []).append(code)

    coders = [c.strip() for c in (co_coders or ["Coder 2"]) if c.strip()]
    overlap_codes: set[str] = set()
    assignments: dict[str, str] = {}

    if len(coders) == 1:
        for st in studies:
            assignments[st["code"]] = coders[0]
    else:
        overlap_label = " & ".join(coders)
        remaining: list[str] = []
        # Quietly stratify overlap across clusters without leaking cluster names to the output sheet
        for _, cl_codes in sorted(by_cluster.items()):
            n_ov = max(1, round(len(cl_codes) * overlap_ratio)) if overlap_ratio > 0 and len(cl_codes) >= 2 else 0
            overlap_codes.update(cl_codes[:n_ov])
            remaining.extend(cl_codes[n_ov:])
        for code in overlap_codes:
            assignments[code] = overlap_label
        for idx, code in enumerate(remaining):
            assignments[code] = coders[idx % len(coders)]

    wb_out = openpyxl.Workbook()
    # Sheet 1: Codebook
    ws_cb = wb_out.active
    ws_cb.title = preset.sheet_codebook
    for row in ws_cb_src.iter_rows(values_only=True):
        ws_cb.append(list(row))
    style_sheet(ws_cb, {1: 42, 2: 42, 3: 44})

    # Sheet 2: Blinded Worksheet (ZERO Coder 1 columns!)
    ws_b = wb_out.create_sheet("Blinded_Coding_Sheet")
    ws_b.append(
        [
            "Study_Code",
            "Authors",
            "Year",
            "Title",
            "Venue",
            "DOI",
            "Assigned_Coder",
            "Paradigm_Cluster",
            "Method_Code",
            "Coder_Notes",
        ]
    )
    for st in studies:
        code = st["code"]
        ws_b.append(
            [
                code,
                st["authors"],
                st["year"],
                st["title"],
                st["venue"],
                st["doi"],
                assignments.get(code, coders[0]),
                "",
                "",
                "",
            ]
        )
        if code in overlap_codes:
            r_idx = ws_b.max_row
            for c_idx in range(1, 11):
                ws_b.cell(r_idx, c_idx).fill = OVERLAP_FILL

    end_row = max(ws_b.max_row, 2)
    if cfg.clusters:
        add_safe_range_dropdown(
            ws_b, "H", 2, end_row, preset.sheet_codebook, "A", 2, 1 + len(cfg.clusters),
            prompt_title="Paradigm Cluster",
        )
    if cfg.method_codes:
        add_safe_range_dropdown(
            ws_b, "I", 2, end_row, preset.sheet_codebook, "B", 2, 1 + len(cfg.method_codes),
            prompt_title="Closed Method Code",
        )
    style_sheet(
        ws_b,
        {1: 12, 2: 26, 3: 8, 4: 44, 5: 24, 6: 24, 7: 22, 8: 38, 9: 36, 10: 32},
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(out_path)

    workload_summary = dict(Counter(assignments.values()))
    return {
        "status": "ok",
        "blinded_workbook_path": str(out_path),
        "total_studies": len(studies),
        "shared_overlap_count": len(overlap_codes),
        "workload_summary": workload_summary,
        "coder1_labels_leaked": False,
    }


def compute_interrater_reliability(
    project_dir: str,
    coder2_workbook_path: str | None = None,
    columns: list[str] | None = None,
) -> dict[str, Any]:
    """Compare Coder 1 (Extraction_Matrix.xlsx) against Coder 2 (blinded workbook) using κ and PABAK."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    preset = cfg.preset()
    matrix_path = cfg.resolve_path(root, "extraction_matrix_file")
    c2_path = Path(coder2_workbook_path) if coder2_workbook_path else cfg.resolve_path(root, "blinded_coder_file")
    if not c2_path.is_absolute():
        c2_path = root / c2_path
    if not c2_path.exists():
        raise FileNotFoundError(f"Coder 2 workbook not found: {c2_path}")

    target_cols = columns or ["Paradigm_Cluster", "Method_Code"]

    wb1 = openpyxl.load_workbook(matrix_path, data_only=True)
    ws1 = wb1[preset.sheet_matrix]
    h1 = [c.value for c in ws1[1]]
    col1 = {h: i for i, h in enumerate(h1) if h}
    p1_records: dict[str, dict[str, str]] = {}
    for row in list(ws1.iter_rows(values_only=True))[1:]:
        if not row or not row[0]:
            continue
        code = str(row[col1["Study_Code"]]).strip()
        p1_records[code] = {
            c: str(row[col1[c]] or "").strip() for c in target_cols if c in col1
        }

    wb2 = openpyxl.load_workbook(c2_path, data_only=True)
    ws2 = wb2["Blinded_Coding_Sheet"] if "Blinded_Coding_Sheet" in wb2.sheetnames else wb2[wb2.sheetnames[-1]]
    h2 = [c.value for c in ws2[1]]
    col2 = {h: i for i, h in enumerate(h2) if h}
    p2_records: dict[str, dict[str, str]] = {}
    for row in list(ws2.iter_rows(values_only=True))[1:]:
        if not row or not row[0]:
            continue
        code = str(row[col2["Study_Code"]]).strip()
        p2_records[code] = {
            c: str(row[col2[c]] or "").strip() for c in target_cols if c in col2
        }

    metrics: dict[str, Any] = {}
    disagreements: list[dict[str, str]] = []
    missing_coder2: list[str] = []

    for c_name in target_cols:
        y1: list[str] = []
        y2: list[str] = []
        for code, p1_vals in sorted(p1_records.items()):
            v1 = p1_vals.get(c_name, "")
            v2 = p2_records.get(code, {}).get(c_name, "")
            if not v2:
                if code not in missing_coder2:
                    missing_coder2.append(code)
                continue
            y1.append(v1)
            y2.append(v2)
            if v1 != v2:
                disagreements.append(
                    {
                        "study_code": code,
                        "dimension": c_name,
                        "coder1": v1,
                        "coder2": v2,
                    }
                )
        metrics[c_name] = cohens_kappa_and_pabak(y1, y2)

    return {
        "status": "ok",
        "total_coder1_studies": len(p1_records),
        "unrated_by_coder2_count": len(missing_coder2),
        "unrated_study_codes": missing_coder2,
        "metrics": metrics,
        "disagreements_count": len(disagreements),
        "disagreements": disagreements,
    }
