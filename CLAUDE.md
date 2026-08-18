## Project
Offerte Vergelijker — Streamlit app that compares supplier PDF offertes against NETTO price lists and the IB budget for building projects.

## Stack
- Python 3.11
- Streamlit (>=1.35,<1.50)
- pandas, openpyxl, pdfplumber, rapidfuzz, streamlit-pdf-viewer, pillow
- Deployed on Azure App Service (Linux)

## Structure
- `IB_met_OfferteVergelijker/Home.py` — landing page, links to each leverancier section
- `IB_met_OfferteVergelijker/common.py` — shared page config, styling and header helpers (use these on every page)
- `IB_met_OfferteVergelijker/pages/1_🟡_Van_Keulen.py` — Van Keulen page: Streamlit UI only
- `IB_met_OfferteVergelijker/pages/2_🧊_Koeling.py` — Koeling page: Streamlit UI only (Koel Installatie + Koel Meubel sub-flows)
- `IB_met_OfferteVergelijker/matching_van_keulen.py` / `matching_koeling.py` — the actual parsing/matching logic for each page, kept Streamlit-free so it's unit-testable (see Verification)
- `IB_met_OfferteVergelijker/data/` — article-number mapping CSVs per leverancier
- `IB_met_OfferteVergelijker/assets/` — logos/icons
- `IB_met_OfferteVergelijker/startup.sh` — Azure App Service startup command
- `tests/` — pytest tests for the `matching_*` modules (repo root)
- `requirements.txt` / `requirements-dev.txt` — runtime / dev (+ruff, +pytest) dependencies (repo root)

## Commands
- Dev: `cd IB_met_OfferteVergelijker && python -m streamlit run Home.py` (http://localhost:8501)
- Install deps: `pip install -r requirements-dev.txt` (or `requirements.txt` for runtime-only)
- Test: `pytest` (from repo root)
- Lint: `ruff check .` (from repo root)
- Deploy: `az webapp up -n offerte-vergelijker --resource-group rg-offerte-app --plan plan-offerte --runtime "PYTHON:3.11"` from the repo root (full setup in `IB_met_OfferteVergelijker/README.md`)

## Verification
No type checker is configured. After every change:
1. Run `ruff check .` and `pytest` from the repo root — fix anything they flag.
2. If the change touches a `pages/*.py` file (Streamlit UI/wiring), also run the app locally and manually exercise the changed page (golden path + edge cases) — pytest only covers the `matching_*` modules, not the Streamlit UI itself.

## Conventions
- Parsing/matching logic lives in a `matching_<section>.py` module (no `streamlit` import, no top-level executing code) so it can be unit-tested; the `pages/*.py` file imports from it and only handles Streamlit UI/wiring. Add new comparison logic to the module, not the page.
- Every page calls `configure_page`, `inject_base_style`, and `jumbo_header`/`back_to_overview` from `common.py` for consistent styling — don't duplicate CSS/layout in a page.
- Matching between offerte and IB budget rows is never a plain key join: try an exact article-number join first, then a `rapidfuzz`-based description fallback. `rapidfuzz` is an optional import guarded by `HAS_RAPIDFUZZ` — keep that guard when touching fuzzy-matching code.
- Domain terms stay in Dutch (offerte, aantal, prijs, omschrijving, leverancier) even though identifiers/code are English — match existing naming rather than translating.
- Row-level data is modeled with `@dataclass` (e.g. `OfferteItem`, `IBItem`), not raw dicts/tuples.
- Supplier-specific quirks (e.g. hardcoded article-number mappings) belong in `IB_met_OfferteVergelijker/data/*.csv`, not inline in page code.
- On every results page, the first `st.tabs(...)` entry is always "📊 Samenvatting" — that's the tab a user should land on right after clicking Analyseren, since Streamlit always shows the first tab passed to `st.tabs` as the default. Other tabs (Alle resultaten, Te controleren/Fuzzy matches, Export) follow after it.

## Don't
- Don't assume offerte and IB budget share an article number — confirm the matching strategy (exact/manual/fuzzy) per leverancier before writing comparison logic.
- Don't add a new page without wiring it into `Home.py` and reusing the `common.py` helpers.
- Don't commit secrets/credentials for Azure — auth is handled via `az login` / Azure AD, not in-repo config.
