"""Project configuration schema and bilingual (EN/ID) localization presets."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import yaml

CONFIG_FILENAME = "scholarflow.yaml"


@dataclass
class LocalizationPreset:
    """Folder and artifact names for a given language preset."""

    language: str
    admin_dir: str
    manuscript_dir: str
    figures_dir: str
    literature_dir: str
    core_dir: str
    unclassified_dir: str
    excluded_dir: str
    methodology_dir: str
    regulation_dir: str
    background_dir: str
    misfiltered_dir: str
    extraction_dir: str
    reading_notes_dir: str
    exports_dir: str
    tools_dir: str
    search_log_file: str
    extraction_matrix_file: str
    blinded_coder_file: str
    protocol_file: str
    manuscript_docx_file: str
    manuscript_md_file: str
    bib_file: str
    ris_file: str
    excluded_log_md: str
    misfiltered_manifest_csv: str
    # Sheet names
    sheet_flow: str
    sheet_candidates: str
    sheet_download_queue: str
    sheet_matrix: str
    sheet_excluded_ft: str
    sheet_codebook: str
    sheet_coder2: str


PRESETS: dict[str, LocalizationPreset] = {
    "en": LocalizationPreset(
        language="en",
        admin_dir="00_Admin",
        manuscript_dir="01_Manuscript",
        figures_dir="01_Manuscript/figures",
        literature_dir="02_Literature",
        core_dir="02_Literature/1_Core",
        unclassified_dir="02_Literature/2_Unclassified",
        excluded_dir="02_Literature/3_Excluded",
        methodology_dir="02_Literature/4_Methodology",
        regulation_dir="02_Literature/5_Policy_Standards",
        background_dir="02_Literature/6_Background",
        misfiltered_dir="02_Literature/archive_misfiltered",
        extraction_dir="03_Extraction_Coding",
        reading_notes_dir="03_Extraction_Coding/Reading_Notes",
        exports_dir="03_Extraction_Coding/Exports",
        tools_dir="04_Tools",
        search_log_file="03_Extraction_Coding/Search_Log.xlsx",
        extraction_matrix_file="03_Extraction_Coding/Extraction_Matrix.xlsx",
        blinded_coder_file="03_Extraction_Coding/Blinded_Coder2_Worksheet.xlsx",
        protocol_file="00_Admin/Review_Protocol.docx",
        manuscript_docx_file="01_Manuscript/Manuscript_Sync.docx",
        manuscript_md_file="01_Manuscript/Manuscript_Sync.md",
        bib_file="04_Tools/references.bib",
        ris_file="04_Tools/references.ris",
        excluded_log_md="03_Extraction_Coding/Reading_Notes/_Excluded_FullText.md",
        misfiltered_manifest_csv="02_Literature/archive_misfiltered/_Manifest_Misfiltered.csv",
        sheet_flow="Flow_Reconciliation",
        sheet_candidates="Candidates",
        sheet_download_queue="Fulltext_Download_Queue",
        sheet_matrix="Extraction_Matrix",
        sheet_excluded_ft="Excluded_FullText",
        sheet_codebook="Codebook",
        sheet_coder2="Coder2_Worksheet",
    ),
    "id": LocalizationPreset(
        language="id",
        admin_dir="00_Admin",
        manuscript_dir="01_Naskah",
        figures_dir="01_Naskah/gambar",
        literature_dir="02_Literatur",
        core_dir="02_Literatur/1_Inti",
        unclassified_dir="02_Literatur/2_Belum_Diklasifikasi",
        excluded_dir="02_Literatur/3_Dikecualikan",
        methodology_dir="02_Literatur/4_Referensi_Metodologi",
        regulation_dir="02_Literatur/5_Regulasi",
        background_dir="02_Literatur/6_Referensi_Latar",
        misfiltered_dir="02_Literatur/arsip_salah_saring_kueri",
        extraction_dir="03_Ekstraksi_Koding",
        reading_notes_dir="03_Ekstraksi_Koding/Catatan_Bacaan",
        exports_dir="03_Ekstraksi_Koding/Ekspor_BasisData",
        tools_dir="04_Tools",
        search_log_file="03_Ekstraksi_Koding/Log_Pencarian.xlsx",
        extraction_matrix_file="03_Ekstraksi_Koding/Matriks_Induk.xlsx",
        blinded_coder_file="03_Ekstraksi_Koding/Lembar_Kerja_Buta_Pengode2.xlsx",
        protocol_file="00_Admin/Protokol_Tinjauan.docx",
        manuscript_docx_file="01_Naskah/Naskah_Tersinkronisasi.docx",
        manuscript_md_file="01_Naskah/Naskah_Tersinkronisasi.md",
        bib_file="04_Tools/Referensi.bib",
        ris_file="04_Tools/Referensi.ris",
        excluded_log_md="03_Ekstraksi_Koding/Catatan_Bacaan/_Dikecualikan_TeksLengkap.md",
        misfiltered_manifest_csv="02_Literatur/arsip_salah_saring_kueri/_Manifest_Salah_Saring.csv",
        sheet_flow="Alur_SPAR4SLR",
        sheet_candidates="Kandidat",
        sheet_download_queue="Unduh_Teks_Lengkap",
        sheet_matrix="Matriks_Induk",
        sheet_excluded_ft="Dikeluarkan_TL",
        sheet_codebook="Buku_Kode_Pengode2",
        sheet_coder2="Lembar_Kerja_Pengode2",
    ),
}


DEFAULT_CLUSTERS = [
    "Cluster A: Market & Statutory Appraisal",
    "Cluster B: Non-Market Economic Valuation",
]

DEFAULT_METHOD_CODES = [
    "M1: Market Comparison / Hedonic",
    "M2: Income / DCF Approach",
    "M3: Cost / Depreciated Replacement Cost",
    "M4: Contingent Valuation (CVM)",
    "M5: Travel Cost Method (TCM)",
    "M6: Choice Experiment (CE)",
    "M7: Spatial Econometrics / GIS-ML",
    "M8: Conceptual / Legal / Qualitative",
]

DEFAULT_EXCLUSION_CODES = [
    "E1: Out of domain / wrong population",
    "E2: No valuation or empirical/methodological core",
    "E3: Non-primary source / editorial / secondary review",
    "E4: Duplicate substance / overlapping dataset",
]


@dataclass
class ProjectConfig:
    """Serializable configuration stored in `<project_root>/scholarflow.yaml`."""

    project_title: str
    authors: list[str] = field(default_factory=lambda: ["Lead Researcher"])
    language: str = "en"  # "en" | "id"
    framework: str = "spar4slr"  # "spar4slr" | "prisma2020" | "empirical"
    research_questions: list[str] = field(default_factory=list)
    clusters: list[str] = field(default_factory=lambda: list(DEFAULT_CLUSTERS))
    method_codes: list[str] = field(default_factory=lambda: list(DEFAULT_METHOD_CODES))
    exclusion_codes: list[str] = field(default_factory=lambda: list(DEFAULT_EXCLUSION_CODES))
    include_ado_tcm: bool = True
    abstract_min_words: int = 150
    abstract_max_words: int = 300
    expected_figures: int = 3
    forbidden_patterns: list[str] = field(
        default_factory=lambda: ["TODO", "TBD", "FIXME", "INSERT CITATION", "CATATAN KERJA"]
    )
    paths: dict[str, str] = field(default_factory=dict)

    def preset(self) -> LocalizationPreset:
        lang = self.language if self.language in PRESETS else "en"
        return PRESETS[lang]

    def resolve_path(self, project_root: str | Path, key: str) -> Path:
        root = Path(project_root)
        if key in self.paths and self.paths[key]:
            return root / self.paths[key]
        preset = self.preset()
        rel = getattr(preset, key)
        return root / rel

    def save(self, project_root: str | Path) -> Path:
        root = Path(project_root)
        cfg_path = root / CONFIG_FILENAME
        data = asdict(self)
        cfg_path.write_text(
            yaml.safe_dump(data, sort_keys=False, allow_unicode=True),
            encoding="utf-8",
        )
        return cfg_path

    @classmethod
    def load(cls, project_root: str | Path) -> ProjectConfig:
        root = Path(project_root)
        cfg_path = root / CONFIG_FILENAME
        if not cfg_path.exists():
            raise FileNotFoundError(
                f"No {CONFIG_FILENAME} found in '{root}'. "
                "Run `init_research_project` first."
            )
        raw: dict[str, Any] = yaml.safe_load(cfg_path.read_text(encoding="utf-8")) or {}
        known = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in raw.items() if k in known}
        return cls(**filtered)
