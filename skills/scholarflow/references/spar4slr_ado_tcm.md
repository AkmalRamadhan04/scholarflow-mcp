# SPAR-4-SLR & ADO–TCM Reference Guide

## 1. Five-Pathway SPAR-4-SLR Accounting Invariants

ScholarFlow enforces exact arithmetic conservation across all three SPAR-4-SLR phases (Paul et al., 2021):

1. **Assembling (Multi-Pathway Identification):**
   $$U_{\text{total}} = P_1 + P_2 + P_3 + P_4 + P_5$$
   - $P_1$: Primary database unique records (e.g., Scopus)
   - $P_2$: Secondary database unique non-duplicate records (OpenAlex, Crossref, SciELO, Google Scholar/PoP)
   - $P_3$: Complementary web / grey literature records
   - $P_4$: Initial researcher seed corpus (non-duplicate)
   - $P_5$: Backward & forward snowballing records (non-duplicate)

2. **Arranging (Screening & Full-Text Retrieval):**
   $$C_{\text{TA}} = U_{\text{total}} - E_{\text{TA}}$$
   $$A_{\text{FT}} = C_{\text{TA}} - R_{\text{unretrieved}}$$
   Every single unretrieved record in $R_{\text{unretrieved}}$ must be explicitly listed in the clickable `Fulltext_Download_Queue` sheet.

3. **Assessing (Full-Text Eligibility & Final Core Corpus):**
   $$N_{\text{core}} = A_{\text{FT}} - E_{\text{FT}}$$
   Every excluded full-text study in $E_{\text{FT}}$ must carry an explicit ordered exclusion code (`E1`–`E4`) in both `Excluded_FullText` and `_Excluded_FullText.md`.

## 2. ADO–TCM Synthesis Structure

- **Antecedents (A):** Drivers, determinants, or structural conditions prompting the phenomenon.
- **Decisions (D):** Analytical choices, policy mechanisms, or valuation approaches adopted.
- **Outcomes (O):** Measured economic, social, spatial, or institutional impacts.
- **Theories (T):** Grounding theoretical lenses.
- **Contexts (C):** Geographic, institutional, legal, and temporal settings.
- **Methods (M):** Specific empirical/analytical designs mapped to closed method codes (`M1–Mn`).
