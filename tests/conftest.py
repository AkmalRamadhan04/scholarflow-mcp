"""Shared pytest fixtures and synthetic PDF helpers (zero private data)."""

from __future__ import annotations

from pathlib import Path

import pymupdf
import pytest


def create_valid_pdf(path: Path, text: str = "Synthetic Academic Paper Content") -> Path:
    """Create a minimal valid 1-page PDF using PyMuPDF."""
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((72, 72), text)
    doc.save(str(path))
    doc.close()
    return path


@pytest.fixture
def synthetic_studies() -> list[dict[str, str]]:
    """Six clean synthetic primary studies spanning 2 paradigm clusters & multiple methods."""
    return [
        {
            "study_code": "S01",
            "authors": "Alvarez, M.; Chen, L.",
            "year": "2021",
            "title": "Spatial Hedonic Pricing of Urban Pocket Parks in Mid-Sized Cities",
            "country_context": "Spain",
            "venue": "Journal of Urban Valuation",
            "doi": "10.1000/demo.2021.001",
            "paradigm_cluster": "Cluster A: Market & Statutory Appraisal",
            "method_code": "M1: Market Comparison / Hedonic",
            "sample_or_object": "1,240 residential sales",
            "key_findings": "Pocket parks generate a 6.4% hedonic price premium within 250 m.",
            "ado_antecedents": "Urban green accessibility",
            "ado_decisions": "Hedonic market capitalization",
            "ado_outcomes": "Higher residential land values",
            "tcm_theories": "Rosen Hedonic Pricing Theory",
            "tcm_contexts": "Mid-sized European cities",
            "tcm_methods_detail": "Spatial lag hedonic regression",
        },
        {
            "study_code": "S02",
            "authors": "Bhattacharya, R.",
            "year": "2022",
            "title": "Contingent Valuation of Heritage Urban Forests under Heat Stress",
            "country_context": "India",
            "venue": "Ecological Economics Demo",
            "doi": "10.1000/demo.2022.002",
            "paradigm_cluster": "Cluster B: Non-Market Economic Valuation",
            "method_code": "M4: Contingent Valuation (CVM)",
            "sample_or_object": "450 urban households",
            "key_findings": "Mean annual WTP for urban forest cooling equals USD 18.50 per household.",
            "ado_antecedents": "Urban heat island stress",
            "ado_decisions": "Household conservation WTP",
            "ado_outcomes": "Quantified non-use welfare value",
            "tcm_theories": "Random Utility Theory",
            "tcm_contexts": "South Asian metropolitan area",
            "tcm_methods_detail": "Double-bounded dichotomous choice CVM",
        },
        {
            "study_code": "S03",
            "authors": "Carvalho, T.; Silva, P.",
            "year": "2023",
            "title": "Depreciated Replacement Cost Models for Municipal Civic Assets",
            "country_context": "Brazil",
            "venue": "Property Appraisal Quarterly",
            "doi": "10.1000/demo.2023.003",
            "paradigm_cluster": "Cluster A: Market & Statutory Appraisal",
            "method_code": "M3: Cost / Depreciated Replacement Cost",
            "sample_or_object": "38 municipal civic facilities",
            "key_findings": "Age-life depreciation adjustments improve public accounting transparency.",
            "ado_antecedents": "Public-sector IPSAS reporting",
            "ado_decisions": "DRC appraisal protocol",
            "ado_outcomes": "Auditable municipal balance sheets",
            "tcm_theories": "Public Asset Valuation Theory",
            "tcm_contexts": "Brazilian municipalities",
            "tcm_methods_detail": "Depreciated replacement cost modeling",
        },
        {
            "study_code": "S04",
            "authors": "Diallo, A.; Mensah, K.; Boateng, E.",
            "year": "2024",
            "title": "Travel Cost Valuation of Coastal Green Corridors",
            "country_context": "Ghana",
            "venue": "Regional Sustainability Studies",
            "doi": "10.1000/demo.2024.004",
            "paradigm_cluster": "Cluster B: Non-Market Economic Valuation",
            "method_code": "M5: Travel Cost Method (TCM)",
            "sample_or_object": "610 corridor visitors",
            "key_findings": "Consumer surplus per visit is USD 9.20 via truncated negative binomial TCM.",
            "ado_antecedents": "Coastal recreation demand",
            "ado_decisions": "Individual trip frequency",
            "ado_outcomes": "Recreational consumer surplus",
            "tcm_theories": "Household Production Theory",
            "tcm_contexts": "West African coastal corridor",
            "tcm_methods_detail": "Zero-truncated negative binomial TCM",
        },
        {
            "study_code": "S05",
            "authors": "Fajar, D.; Kusuma, N.",
            "year": "2024",
            "title": "Explainable GeoAI Mass Appraisal of Green-Adjacent Land Zones",
            "country_context": "Indonesia",
            "venue": "Spatial Economics Review",
            "doi": "10.1000/demo.2024.006",
            "paradigm_cluster": "Cluster A: Market & Statutory Appraisal",
            "method_code": "M7: Spatial Econometrics / GIS-ML",
            "sample_or_object": "2,100 statutory land zones",
            "key_findings": "Spatial block CV + GeoShapley reveals non-linear green-space capitalization.",
            "ado_antecedents": "Statutory land value zoning",
            "ado_decisions": "Explainable spatial ML appraisal",
            "ado_outcomes": "Reduced assessment regressivity",
            "tcm_theories": "Spatial Heterogeneity & XAI",
            "tcm_contexts": "Southeast Asian megacity",
            "tcm_methods_detail": "XGBoost + GeoShapley + MGWR",
        },
        {
            "study_code": "S06",
            "authors": "Gomez, S.; Park, J.",
            "year": "2025",
            "title": "Choice Experiments for Climate-Resilient Public Plazas",
            "country_context": "South Korea",
            "venue": "Environmental Valuation Journal",
            "doi": "10.1000/demo.2025.007",
            "paradigm_cluster": "Cluster B: Non-Market Economic Valuation",
            "method_code": "M6: Choice Experiment (CE)",
            "sample_or_object": "520 survey respondents",
            "key_findings": "Permeable pavement and shade canopy drive the highest marginal WTP.",
            "ado_antecedents": "Pluvial flood & heat resilience",
            "ado_decisions": "Discrete choice attribute trade-offs",
            "ado_outcomes": "Prioritized green plaza investments",
            "tcm_theories": "Lancaster Characteristics Theory",
            "tcm_contexts": "East Asian dense urban district",
            "tcm_methods_detail": "Mixed logit discrete choice model",
        },
    ]
