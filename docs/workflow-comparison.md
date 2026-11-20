# Workflow Comparison: Standard AI Research Skills vs. ScholarFlow MCP

This document outlines how **ScholarFlow MCP** extends conventional prompt-based academic research skills (such as `deep-research`, `sr-screener`, and `academic-pipeline`) with deterministic tooling forged during real multi-database systematic literature reviews.

| Stage | Standard Prompt-Based Academic Skills | ScholarFlow MCP Addition |
| :--- | :--- | :--- |
| **0. Workspace Setup** | Ad-hoc markdown files in working directory | Structured 5-folder research archive (`00_Admin`–`04_Tools`) with `scholarflow.yaml` and bilingual (`EN`/`ID`) presets |
| **1. Identification & Reconciliation** | Basic search notes or PRISMA-only counts | Multi-DB CSV/RIS/BibTeX deduplication across **5 explicit SPAR-4-SLR pathways** + strict conservation equations |
| **2. Full-Text Acquisition** | Manual download or best-effort URL lookup | Ethical OpenAlex + `citation_pdf_url`/OJS OA downloader (strict CAPTCHA refusal) + clickable Excel **manual queue** |
| **3. Corpus File Governance** | Not tracked on disk | Non-destructive physical routing + misfiltered CSV manifest + `PyMuPDF` corruption audit + 1-to-1 reading notes |
| **4. Data Extraction & Coding** | Explicitly out of scope in screening skills | Closed-taxonomy `Extraction_Matrix.xlsx` + **ADO–TCM** synthesis columns + Excel-safe cell-range dropdowns |
| **5. Inter-Rater Reliability** | Dual-LLM screening only | **100%-blinded human co-coder export** with cluster-stratified overlap + **Cohen's Kappa ($\kappa$) & PABAK** |
| **6. Bibliography Integrity** | Markdown references or unverified BibTeX | Optional live Crossref/OpenAlex DOI check + clean `.bib` & `.ris` (0 `&`-joined `AU  -` bugs) |
| **7. Manuscript & Figures** | Manually transcribed tables & statistics | **Single Source of Truth**: 300 DPI charts + synchronized DOCX/MD compiled directly from Excel |
| **8. Pre-Submission Audit** | LLM self-check | **20+ deterministic 4-way checks** returning machine-readable `PASS`/`FAIL` JSON reports |
