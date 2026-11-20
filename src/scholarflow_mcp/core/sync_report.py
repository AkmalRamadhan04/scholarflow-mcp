"""Stage 7: Data-driven 300 DPI figures & Single-Source-of-Truth DOCX/Markdown manuscript synchronization."""

from __future__ import annotations

from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any

import matplotlib
import openpyxl
from docx import Document
from docx.shared import Inches

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

from scholarflow_mcp.config import ProjectConfig
from scholarflow_mcp.core.bibliography import split_authors_clean


def _load_matrix_and_flow(root: Path, cfg: ProjectConfig) -> tuple[list[dict[str, Any]], dict[str, int]]:
    preset = cfg.preset()
    wb_m = openpyxl.load_workbook(cfg.resolve_path(root, "extraction_matrix_file"), data_only=True)
    ws_m = wb_m[preset.sheet_matrix]
    headers = [c.value for c in ws_m[1]]
    col = {h: i for i, h in enumerate(headers) if h}

    studies: list[dict[str, Any]] = []
    for row in list(ws_m.iter_rows(values_only=True))[1:]:
        if not row or not row[0]:
            continue
        studies.append({h: (row[idx] if idx < len(row) else "") for h, idx in col.items()})

    wb_l = openpyxl.load_workbook(cfg.resolve_path(root, "search_log_file"), data_only=True)
    ws_f = wb_l[preset.sheet_flow]
    flow = {
        str(row[0]): int(row[2] or 0)
        for row in list(ws_f.iter_rows(values_only=True))[1:]
        if row and row[0]
    }
    return studies, flow


def _draw_flow_diagram(out_path: Path, flow: dict[str, int], framework: str) -> None:
    fig, ax = plt.subplots(figsize=(9.0, 6.2), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis("off")

    title = (
        "SPAR-4-SLR 5-Pathway Selection Flow"
        if framework == "spar4slr"
        else "PRISMA 2020 Systematic Selection Flow"
    )
    ax.text(5.0, 9.6, title, ha="center", va="center", fontsize=13, fontweight="bold", color="#1F4E78")

    boxes = [
        (
            0.6,
            7.6,
            8.8,
            1.4,
            "#D9E1F2",
            f"ASSEMBLING — Multi-Pathway Identification (Unique Total = {flow.get('total_unique_screened', 0)})\n"
            f"P1 Primary DB: {flow.get('pathway_1_primary_db', 0)} | "
            f"P2 Secondary DB: {flow.get('pathway_2_secondary_db', 0)} | "
            f"P3 Web: {flow.get('pathway_3_web_complement', 0)} | "
            f"P4 Seed: {flow.get('pathway_4_team_seed', 0)} | "
            f"P5 Snowballing: {flow.get('pathway_5_snowballing', 0)}",
        ),
        (
            0.6,
            5.5,
            5.0,
            1.2,
            "#EBF1FA",
            f"ARRANGING — Passed Title/Abstract\nCandidates = {flow.get('candidates_passing_ta', 0)}",
        ),
        (
            6.1,
            5.5,
            3.3,
            1.2,
            "#FCE4D6",
            f"Excluded at T/A\nn = {flow.get('excluded_title_abstract', 0)}",
        ),
        (
            0.6,
            3.3,
            5.0,
            1.2,
            "#EBF1FA",
            f"ASSESSING — Full-Text Assessed\nAssessed = {flow.get('fulltext_assessed', 0)}",
        ),
        (
            6.1,
            3.3,
            3.3,
            1.2,
            "#FCE4D6",
            f"Not Retrieved\nn = {flow.get('unretrieved_fulltext', 0)}",
        ),
        (
            0.6,
            1.0,
            5.0,
            1.3,
            "#E2EFDA",
            f"FINAL LOCKED CORE CORPUS\nN = {flow.get('final_core_corpus', 0)} Studies",
        ),
        (
            6.1,
            1.0,
            3.3,
            1.3,
            "#FCE4D6",
            f"Excluded at Full-Text\nn = {flow.get('excluded_fulltext', 0)}",
        ),
    ]

    for x, y, w, h, fc, text in boxes:
        patch = FancyBboxPatch(
            (x, y),
            w,
            h,
            boxstyle="round,pad=0.08,rounding_size=0.15",
            facecolor=fc,
            edgecolor="#1F4E78",
            linewidth=1.2,
        )
        ax.add_patch(patch)
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=8.8, color="#1A1A1A")

    arrow_kw = {"arrowstyle": "->", "color": "#1F4E78", "lw": 1.4}
    ax.annotate("", xy=(3.1, 6.75), xytext=(3.1, 7.55), arrowprops=arrow_kw)
    ax.annotate("", xy=(3.1, 4.55), xytext=(3.1, 5.45), arrowprops=arrow_kw)
    ax.annotate("", xy=(3.1, 2.35), xytext=(3.1, 3.25), arrowprops=arrow_kw)
    ax.annotate("", xy=(6.05, 6.1), xytext=(5.65, 6.1), arrowprops=arrow_kw)
    ax.annotate("", xy=(6.05, 3.9), xytext=(5.65, 3.9), arrowprops=arrow_kw)
    ax.annotate("", xy=(6.05, 1.65), xytext=(5.65, 1.65), arrowprops=arrow_kw)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _draw_temporal_cluster_chart(out_path: Path, studies: list[dict[str, Any]]) -> None:
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10.0, 4.2), dpi=300)

    years = [int(str(s.get("Year", ""))[:4]) for s in studies if str(s.get("Year", ""))[:4].isdigit()]
    yr_counts = Counter(years)
    if yr_counts:
        xs = sorted(yr_counts)
        ys = [yr_counts[x] for x in xs]
        ax1.bar([str(x) for x in xs], ys, color="#1F4E78", edgecolor="white")
    ax1.set_title("Publication Year Distribution", fontsize=11, fontweight="bold", color="#1F4E78")
    ax1.set_xlabel("Year")
    ax1.set_ylabel("Studies (n)")
    ax1.grid(axis="y", linestyle="--", alpha=0.4)

    clusters = Counter(str(s.get("Paradigm_Cluster") or "Unassigned") for s in studies)
    if clusters:
        labels = [lbl.split(":")[0] for lbl in clusters]
        vals = list(clusters.values())
        colors = ["#1F4E78", "#2E75B6", "#70AD47", "#ED7D31", "#5B9BD5"][: len(labels)]
        ax2.barh(labels, vals, color=colors)
        for i, v in enumerate(vals):
            ax2.text(v + 0.05, i, str(v), va="center", fontsize=9, fontweight="bold")
    ax2.set_title("Paradigm Cluster Balance", fontsize=11, fontweight="bold", color="#1F4E78")
    ax2.set_xlabel("Studies (n)")
    ax2.grid(axis="x", linestyle="--", alpha=0.4)

    fig.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _draw_method_taxonomy_chart(out_path: Path, studies: list[dict[str, Any]]) -> None:
    fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=300)
    methods = Counter(str(s.get("Method_Code") or "Unassigned") for s in studies)
    if methods:
        items = sorted(methods.items(), key=lambda kv: (-kv[1], kv[0]))
        labels = [k for k, _ in items]
        vals = [v for _, v in items]
        ax.barh(labels[::-1], vals[::-1], color="#2E75B6", edgecolor="white")
        for i, v in enumerate(vals[::-1]):
            ax.text(v + 0.05, i, str(v), va="center", fontsize=9, fontweight="bold")
    ax.set_title("Closed-Taxonomy Method Distribution", fontsize=11, fontweight="bold", color="#1F4E78")
    ax.set_xlabel("Studies (n)")
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    fig.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def generate_slr_figures(project_dir: str) -> dict[str, Any]:
    """Generate 3 publication-ready 300 DPI PNG figures directly from Search_Log.xlsx & Extraction_Matrix.xlsx."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    fig_dir = cfg.resolve_path(root, "figures_dir")
    studies, flow = _load_matrix_and_flow(root, cfg)

    f1 = fig_dir / "Fig1_Flow_Diagram.png"
    f2 = fig_dir / "Fig2_Temporal_Cluster_Distribution.png"
    f3 = fig_dir / "Fig3_Method_Taxonomy_Distribution.png"

    _draw_flow_diagram(f1, flow, cfg.framework)
    _draw_temporal_cluster_chart(f2, studies)
    _draw_method_taxonomy_chart(f3, studies)

    return {
        "status": "ok",
        "figures_dir": str(fig_dir),
        "dpi": 300,
        "figures": [str(f1), str(f2), str(f3)],
    }


def _format_apa7_reference(st: dict[str, Any]) -> str:
    authors = split_authors_clean(str(st.get("Authors") or "Anonymous"))
    if len(authors) == 1:
        auth_str = authors[0]
    elif len(authors) == 2:
        auth_str = f"{authors[0]}, & {authors[1]}"
    else:
        auth_str = ", ".join(authors[:-1]) + f", & {authors[-1]}"
    year = str(st.get("Year") or "n.d.")[:4]
    title = str(st.get("Title") or "Untitled").rstrip(".")
    venue = str(st.get("Venue") or "Unpublished Manuscript").strip()
    doi = str(st.get("DOI") or "").strip()
    doi_part = f" https://doi.org/{doi}" if doi else ""
    return f"{auth_str} ({year}). {title}. {venue}.{doi_part}"


def _short_citation(st: dict[str, Any]) -> str:
    authors = split_authors_clean(str(st.get("Authors") or "Anonymous"))
    first_surname = authors[0].split(",")[0].strip().split()[-1] if authors else "Anonymous"
    year = str(st.get("Year") or "n.d.")[:4]
    if len(authors) == 1:
        return f"{first_surname} ({year})"
    if len(authors) == 2:
        second_surname = authors[1].split(",")[0].strip().split()[-1]
        return f"{first_surname} & {second_surname} ({year})"
    return f"{first_surname} et al. ({year})"


def sync_manuscript_snapshot(
    project_dir: str,
    custom_abstract: str | None = None,
    regenerate_figures: bool = True,
    sync_protocol: bool = True,
) -> dict[str, Any]:
    """Single-Source-of-Truth builder: compile synchronized DOCX + Markdown manuscript directly from Excel."""
    root = Path(project_dir).resolve()
    cfg = ProjectConfig.load(root)
    studies, flow = _load_matrix_and_flow(root, cfg)

    if regenerate_figures:
        fig_info = generate_slr_figures(str(root))
        fig_paths = [Path(p) for p in fig_info["figures"]]
    else:
        fig_dir = cfg.resolve_path(root, "figures_dir")
        fig_paths = [
            fig_dir / "Fig1_Flow_Diagram.png",
            fig_dir / "Fig2_Temporal_Cluster_Distribution.png",
            fig_dir / "Fig3_Method_Taxonomy_Distribution.png",
        ]

    n_core = len(studies)
    clusters = Counter(str(s.get("Paradigm_Cluster") or "Unassigned") for s in studies)
    methods = Counter(str(s.get("Method_Code") or "Unassigned") for s in studies)
    countries = sorted({str(s.get("Country_Context") or "").strip() for s in studies if s.get("Country_Context")})
    years = [int(str(s.get("Year", ""))[:4]) for s in studies if str(s.get("Year", ""))[:4].isdigit()]
    yr_span = f"{min(years)}–{max(years)}" if years else "n.d."

    default_abstract = (
        f"Purpose: This systematic literature review investigates {cfg.project_title} using the "
        f"{cfg.framework.upper()} protocol to establish a transparent, auditable, and fully reproducible "
        f"evidence synthesis across multiple scholarly databases and complementary retrieval pathways. "
        f"Design/methodology/approach: A structured five-pathway search strategy assembled "
        f"{flow.get('total_unique_screened', n_core)} unique records across primary databases, secondary "
        f"indexes, complementary web sources, seed collections, and forward-backward snowballing. Following "
        f"title-abstract screening ({flow.get('candidates_passing_ta', n_core)} candidates), full-text "
        f"acquisition ({flow.get('fulltext_assessed', n_core)} assessed), and closed-taxonomy eligibility "
        f"evaluation, a locked core corpus of N = {n_core} primary studies spanning {yr_span} across "
        f"{len(countries) or 1} geographical contexts was extracted into a single-source-of-truth matrix "
        f"with independent coder verification and deterministic four-way integrity auditing. "
        f"Findings: The locked corpus spans {len(clusters)} paradigm clusters and {len(methods)} closed "
        f"methodological categories, revealing complementary theoretical mechanisms, empirical trade-offs, "
        f"and persistent methodological gaps documented through structured ADO–TCM synthesis. "
        f"Originality/value: By coupling strict multi-pathway flow reconciliation, blinded human reliability "
        f"testing, data-driven manuscript generation, and automated four-way integrity verification, this "
        f"review delivers an auditable benchmark that eliminates transcription drift between screening logs, "
        f"coding matrices, bibliographies, and manuscript tables."
    )
    abstract_text = custom_abstract.strip() if custom_abstract else default_abstract

    # Build prose body explicitly citing every locked study and referencing all 3 figures
    cite_clauses = [f"{_short_citation(st)} [{st['Study_Code']}]" for st in studies]
    cluster_summary_lines = [
        f"- **{cl}**: {cnt} studies ({cnt / n_core * 100:.1f}%)" if n_core else f"- **{cl}**: 0"
        for cl, cnt in clusters.most_common()
    ]
    method_summary_lines = [
        f"- **{mc}**: {cnt} studies" for mc, cnt in methods.most_common()
    ]
    apa_refs = sorted(_format_apa7_reference(st) for st in studies)

    # 1. Write DOCX
    docx_path = cfg.resolve_path(root, "manuscript_docx_file")
    doc = Document()
    doc.add_heading(cfg.project_title, level=0)
    doc.add_paragraph(f"Authors: {', '.join(cfg.authors)} | Synced: {date.today().isoformat()}")

    doc.add_heading("ABSTRACT", level=1)
    doc.add_paragraph(abstract_text)
    doc.add_paragraph(f"Keywords: {cfg.framework.upper()}; systematic literature review; reproducibility; ADO-TCM")

    doc.add_heading("1. Introduction", level=1)
    doc.add_paragraph(
        f"This manuscript provides a data-synchronized synthesis of {n_core} locked primary studies "
        f"({yr_span}) assembled under the {cfg.framework.upper()} protocol."
    )

    doc.add_heading("2. Methodology and Selection Flow", level=1)
    doc.add_paragraph(
        f"As illustrated in Figure 1, multi-pathway assembling yielded {flow.get('total_unique_screened', 0)} "
        f"unique records, resulting in {flow.get('candidates_passing_ta', 0)} title/abstract candidates, "
        f"{flow.get('fulltext_assessed', 0)} assessed full-text studies, and a final locked core corpus of "
        f"N = {n_core} studies."
    )
    if fig_paths[0].exists():
        doc.add_picture(str(fig_paths[0]), width=Inches(5.8))
        doc.add_paragraph("Figure 1. Systematic selection and five-pathway reconciliation flow (300 DPI).")

    doc.add_heading("3. Synthesis of Findings", level=1)
    doc.add_paragraph(
        "The locked primary corpus encompasses "
        + (", ".join(cite_clauses) if cite_clauses else "no studies yet")
        + ". Figure 2 summarizes the temporal evolution and paradigm cluster balance across the corpus."
    )
    for line in cluster_summary_lines:
        doc.add_paragraph(line.replace("**", ""), style="List Bullet")
    if fig_paths[1].exists():
        doc.add_picture(str(fig_paths[1]), width=Inches(5.8))
        doc.add_paragraph("Figure 2. Temporal distribution and paradigm cluster balance (300 DPI).")

    doc.add_paragraph(
        "Figure 3 reports the distribution of closed-taxonomy methodological approaches extracted from "
        "each verified full-text study."
    )
    for line in method_summary_lines:
        doc.add_paragraph(line.replace("**", ""), style="List Bullet")
    if fig_paths[2].exists():
        doc.add_picture(str(fig_paths[2]), width=Inches(5.8))
        doc.add_paragraph("Figure 3. Closed-taxonomy methodological distribution (300 DPI).")

    doc.add_heading("REFERENCES", level=1)
    for ref in apa_refs:
        doc.add_paragraph(ref)

    doc.add_heading("APPENDIX 1. Locked Core Corpus Matrix", level=1)
    tbl = doc.add_table(rows=1, cols=6)
    hdr = tbl.rows[0].cells
    for idx, label in enumerate(["No", "Study_Code", "Authors", "Year", "Paradigm_Cluster", "Method_Code"]):
        hdr[idx].text = label
    for idx, st in enumerate(studies, start=1):
        row_cells = tbl.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(st.get("Study_Code", ""))
        row_cells[2].text = str(st.get("Authors", ""))
        row_cells[3].text = str(st.get("Year", ""))
        row_cells[4].text = str(st.get("Paradigm_Cluster", ""))
        row_cells[5].text = str(st.get("Method_Code", ""))

    docx_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(docx_path)

    # 2. Write synchronized Markdown mirror
    md_path = cfg.resolve_path(root, "manuscript_md_file")
    md_lines = [
        f"# {cfg.project_title}",
        "",
        f"**Authors:** {', '.join(cfg.authors)}  ",
        f"**Synchronized:** {date.today().isoformat()} | **Locked Corpus:** `N = {n_core}` ({yr_span})",
        "",
        "## ABSTRACT",
        abstract_text,
        "",
        f"**Keywords:** {cfg.framework.upper()}; systematic literature review; reproducibility; ADO-TCM",
        "",
        "## 1. Introduction",
        f"This manuscript provides a data-synchronized synthesis of `{n_core}` locked primary studies ({yr_span}).",
        "",
        "## 2. Methodology and Selection Flow",
        f"As illustrated in Figure 1, `{flow.get('total_unique_screened', 0)}` unique records narrowed to `N = {n_core}` locked core studies.",
        "",
        "![Figure 1](figures/Fig1_Flow_Diagram.png)",
        "",
        "## 3. Synthesis of Findings",
        "Primary studies included: " + (", ".join(cite_clauses) if cite_clauses else "None"),
        "",
        *cluster_summary_lines,
        "",
        "![Figure 2](figures/Fig2_Temporal_Cluster_Distribution.png)",
        "",
        *method_summary_lines,
        "",
        "![Figure 3](figures/Fig3_Method_Taxonomy_Distribution.png)",
        "",
        "## REFERENCES",
        *[f"- {ref}" for ref in apa_refs],
        "",
        "## APPENDIX 1. Locked Core Corpus Matrix",
        "| No | Study_Code | Authors | Year | Paradigm_Cluster | Method_Code |",
        "|---:|---|---|---:|---|---|",
        *[
            f"| {i} | {s.get('Study_Code', '')} | {s.get('Authors', '')} | {s.get('Year', '')} | {s.get('Paradigm_Cluster', '')} | {s.get('Method_Code', '')} |"
            for i, s in enumerate(studies, start=1)
        ],
        "",
    ]
    md_path.write_text("\n".join(md_lines), encoding="utf-8")

    # 3. Sync locked N into Review_Protocol.docx
    if sync_protocol:
        prot_path = cfg.resolve_path(root, "protocol_file")
        if prot_path.exists():
            pdoc = Document(prot_path)
            updated = False
            for p in pdoc.paragraphs:
                if "Locked Core Corpus:" in p.text or "Korpus Inti Terkunci:" in p.text:
                    p.text = f"Locked Core Corpus: N = {n_core} studies ({yr_span})."
                    updated = True
                    break
            if not updated:
                pdoc.add_paragraph(f"Locked Core Corpus: N = {n_core} studies ({yr_span}).")
            pdoc.save(prot_path)

    return {
        "status": "ok",
        "manuscript_docx": str(docx_path),
        "manuscript_md": str(md_path),
        "locked_corpus_n": n_core,
        "abstract_words": len(abstract_text.split()),
        "references_count": len(apa_refs),
        "figures_embedded": sum(1 for p in fig_paths if p.exists()),
    }
