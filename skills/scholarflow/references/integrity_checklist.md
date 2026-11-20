# 4-Way Deterministic Integrity Audit Contract

`run_four_way_integrity_audit` verifies four coupled layers simultaneously:

| Direction | Audit Scope | Blocking Failure Examples |
| :--- | :--- | :--- |
| **Direction 1: Physical Corpus** | Staging folder, `1_Core`, `3_Excluded`, `6_Background`, `archive_misfiltered` | Unrouted file in `2_Unclassified`; corrupt/0-page PDF failing `fitz.open()`; core count = 0 |
| **Direction 2: Search Log** | `Flow_Reconciliation` & `Fulltext_Download_Queue` | Broken 5-pathway sum; `Cand_TA - Unretrieved != Assessed`; missing clickable hyperlink on queue |
| **Direction 3: Extraction Matrix & Notes** | `Extraction_Matrix`, `Excluded_FullText`, `Coder2_Worksheet`, `Reading_Notes/` | Duplicate study code; 1-to-1 filename mismatch; empty closed taxonomy; missing reading note; inline Excel DataValidation `>255` chars |
| **Direction 4: Manuscript, Bib/RIS & Protocol** | `Manuscript_Sync.docx`, `.bib`, `.ris`, `Review_Protocol.docx` | Abstract outside word limit; missing 300 DPI figure embed/reference; appendix table drift; forbidden placeholder (`TODO`, legacy errors); uncited reference in prose; `&` inside RIS `AU  -` line; protocol `N` out of sync |

## AI Transparency & Authorship Disclosure

AI agents assist with deterministic scripting, metadata verification, and audit logging. Human authors remain solely accountable for protocol decisions, full-text eligibility judgments, blinded reliability consensus, and substantive scientific claims.
