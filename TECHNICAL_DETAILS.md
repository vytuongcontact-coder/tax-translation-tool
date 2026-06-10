# Technical Documentation: Auto Glossary Extractor (V2)

This document provides a detailed breakdown of the architecture, technical stack, core algorithms, and general applicability of the **Auto Glossary Extractor** platform.

---

## 1. System Architecture

The application is built on a **stateless, serverless-friendly architecture** designed to run on scalable cloud environments (like Vercel) without depending on a persistent filesystem.

```
                    ┌────────────────────────┐
                    │  Frontend Browser UI   │
                    │   (HTML5, CSS3, JS)    │
                    └───────────┬────────────┘
                                │
          1. POST /upload       │      3. POST /format
          (File upload)         │      (Glossary JSON)
                                ▼
         ┌──────────────────────────────────────────┐
         │              Flask Backend               │
         │         (app.py / Vercel Serverless)     │
         └──────────────┬────────────────────┬──────┘
                        │                    │
                        ▼                    ▼
             ┌─────────────────────┐┌─────────────────┐
             │  core/extractor.py  ││core/formatter.py│
             │   (PyMuPDF/docx)    ││(Pandas/openpyxl)│
             └─────────────────────┘└─────────────────┘
```

### Key Architectural Pillars:
* **Stateless Execution**: The backend does not write permanent files to disk. Uploads are processed in-memory or via temporary files (`tempfile`) that are deleted immediately after the request.
* **In-Memory Formatting**: The Excel output is compiled directly into a memory stream (`io.BytesIO`) and streamed to the user's browser, eliminating file storage needs.
* **Client-Side State Management**: The browser maintains the active dictionary terms state (`glossaryTerms`), letting the user review, add, edit, or delete items before triggering final export.

---

## 2. Core Technical Components

### A. Document Parsing & Text Extraction (`core/extractor.py`)
The system parses two main formats (.DOCX and .PDF) and merges them into a normalized string for matching:
* **Word Documents (`.docx`)**: Utilizes `python-docx` to read all paragraphs and table cells, normalizing mid-word line breaks (e.g. converting `Arm's \n Length` into `Arm's Length`).
* **PDF Documents (`.pdf`)**: Utilizes PyMuPDF (`fitz`) to extract text layout-by-layout, handling cross-page word wraps.

### B. High-Speed Greedy Regex Keyword Matching (`core/extractor.py`)
This is the core algorithm of the glossary scanner:
1. **Dictionary Compiling**: Loads `core/tax_dictionary.json` (a curated tax dictionary of hundreds of terms).
2. **Greedy Sort**: Sorts all terms by length in descending order. This ensures that compound terms are evaluated before individual words. For example, `"transfer pricing policy"` is matched first, avoiding matching it as `"transfer pricing"` + `"policy"`.
3. **Single-Pass Regular Expression**: Escapes special characters and joins all terms with the logical OR `|` operator, enclosed by word boundaries (`\b`):
   $$\text{Pattern} = \backslash\text{b}(\text{term}_1|\text{term}_2|\text{term}_3|\dots)\backslash\text{b}$$
   This pattern is compiled with `re.IGNORECASE` for high-speed scanning of the document body.
4. **Case-Insensitive Deduplication**: Matches are returned preserving their original casing in the document, and then deduplicated case-insensitively.

### C. Case Variant Generation & Expansion (`core/formatter.py`)
To maximize the usefulness of the glossary in computer-assisted translation (CAT) tools (such as SDL Trados, memoQ, or Phrase), the export engine expands each term into 5 casing variants and 1 acronym variant:
* **Lowercase**: `transfer pricing`
* **Uppercase**: `TRANSFER PRICING`
* **Capitalized**: `Transfer pricing`
* **Title Case**: `Transfer Pricing`
* **Apostrophe Title Correction**: `Arm's Length` (fixing Python's default `.title()` behavior which produces `Arm'S`)
* **Acronym**: Generates abbreviation variations (e.g. `TRANSFER PRICING GLOBAL MASTER FILE` $\rightarrow$ `TPGMF`) for terms with $\ge 2$ words.

### D. Bi-directional Sync Scrolling & Live Previews (`templates/index.html`)
The frontend contains high-fidelity javascript linkages to sync the document preview and terms list:
* **Quotes Normalization**: To prevent quote mismatching (curly quotes `’` in the PDF vs straight quotes `'` in the dictionary), the text is normalized to straight quotes (`'`) on both sides.
* **Preview ➔ Editor Link**: Clicking a highlighted term in the preview panel searches the browser's glossary array, identifies the index, scrolls the editor row into view, and flashes the border.
* **Editor ➔ Preview Link**: Clicking the search icon next to an editor input triggers a DOM query inside the preview, scrolls the first matched occurrence into center, and highlights it.

---

## 3. Technology Stack

| Layer | Technology | Purpose |
| :--- | :--- | :--- |
| **Backend Framework** | `Flask (Python 3.9+)` | Minimal, high-speed serverless routing |
| **Document Reader** | `PyMuPDF (fitz)` & `python-docx` | Layout-level text extraction from PDF and Word documents |
| **Data & Spreadsheet** | `Pandas` & `openpyxl` | Excel generation, case variant mapping, and sorting |
| **Frontend Structure** | `HTML5` & `JavaScript (ES6)` | Direct DOM manipulation, stateless JSON parsing, and layout syncing |
| **Styling & Theme** | `CSS3 (Vanilla)` | Responsive media queries, KPMG branding colors, and glassmorphic inputs |
| **Cloud Hosting** | `Vercel` | Serverless Python function runtime |

---

## 4. General Applicability (Generalization)

Although this application is customized with a **tax dictionary** (`core/tax_dictionary.json`) and **KPMG Visual Identity**, the core engineering is generic and can be applied to any industry glossary extraction pipeline:

### 1. Dictionary Swap (Any Domain)
By replacing the `core/tax_dictionary.json` file, the tool can instantly scan for:
* **Medical / Pharmaceutical**: Scan clinical trials for drug names, side effects, and symptoms.
* **Legal / Contracts**: Scan legal agreements for contract clauses, corporate names, and compliance terms.
* **Engineering**: Scan construction specifications for materials, standards, and equipment codes.

### 2. Multi-Language Support
While the default UI and excel columns use `Source` (English) and `Target` (Vietnamese), the translation input boxes can accept any Unicode characters (Chinese, Japanese, Arabic, Russian, etc.) and write them successfully to standard UTF-8 Excel formats.

### 3. CAT Tool Integration
The generated spreadsheet format (containing parallel columns of expanded case variants) is designed to be directly imported into translation databases:
* Can be converted to `.tbx` (TermBase eXchange) format.
* Can be imported directly into Trados MultiTerm or memoQ Termbase, saving translators hours of manual copy-pasting.
