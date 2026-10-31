# Aqua Cloud scraper and Word handbook generator

Run one command to collect public Aqua Cloud product pages and documentation and create a detailed `.docx` handbook. No OpenAI key, paid API or Aqua account is required. Python 3.10 or later is required.

The product tour was inspected on 31 October 2026. It contained meaningful server-rendered HTML; treating it as an empty SPA would have been an incorrect starting assumption. The tool uses HTTP first, supports GitBook Markdown responses and can render thin pages with Playwright. It does not guess API URLs.

## Quick start on Windows

Open PowerShell or Command Prompt in this extracted folder:

```powershell
py -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m playwright install chromium
.venv\Scripts\python aqua_scraper.py
```

After this first setup, double-click `run_windows.bat` to run again. It uses the existing virtual environment and does not install anything.

## Quick start on macOS or Linux

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m playwright install chromium
.venv/bin/python aqua_scraper.py
```

For later runs use `bash run_mac_linux.sh`. On a minimal Linux host Chromium may also require system libraries; use Playwright's documented `install-deps chromium` command if its launch error identifies missing dependencies.

## What you get

| File in output | Purpose |
| --- | --- |
| Aqua_Cloud_Handbook.docx | Detailed readable handbook with 14 topic chapters, original source sections and clickable source links |
| aqua_data.json | Collected text blocks, metadata, URLs, content hashes, missing pages, pending URLs and browser request metadata |
| scrape.log | Progress and problems from this and any previous run in the same output folder |

Markdown documentation and linked llms.txt indexes are supported. Indexes are used for discovery rather than printed as product chapters.

The document covers product overview; requirements and traceability; test cases and data; planning and execution; defects and Capture; automation; AI; integrations and APIs; agile collaboration; reports and history; administration and security; migration; pricing; and additional documentation. Chapters without collected material are explicitly marked. The collector does not claim to have accessed the private application or all possible documentation.

Each chapter starts with a short plain-English explanation of the concept. The detailed body then preserves collected source wording, headings and table rows. This is a source-organised handbook, not an AI-written narrative or an independent product review. It can be long, and some vendor repetition can remain. Use the Word Navigation Pane to move between chapters. There is no automatic page-numbered table of contents.

## Choose your scope

The default run attempts up to 50 URLs and follows links to depth 2, starting with the product tour, official documentation, product FAQ and pricing. Failures also count toward the limit so a blocked site cannot cause an unbounded run. The queue is breadth-first. An extensive documentation site can require a separate documentation-focused run.

```bash
# Quick check
python aqua_scraper.py --max-pages 6 --output output_quick

# More detailed product and documentation collection
python aqua_scraper.py --max-pages 100 --max-depth 3 --output output_detailed

# Documentation-focused collection
python aqua_scraper.py --seed https://docs.aqua-cloud.io/documentation --max-pages 100 --max-depth 3 --output output_docs

# Only the product tour, with no linked pages
python aqua_scraper.py --seed https://aqua-cloud.io/product-tour/ --max-pages 1 --max-depth 0

# Force JavaScript rendering
python aqua_scraper.py --mode browser --max-pages 10 --output output_browser

# Direct HTTP only; Chromium is not needed
python aqua_scraper.py --mode http --max-pages 30 --output output_http
```

In these examples `python` means the interpreter inside your virtual environment. Use the explicit `.venv` path from setup if the environment is not activated.

The allowed hosts are `aqua-cloud.io`, `www.aqua-cloud.io` and `docs.aqua-cloud.io`. Documentation URLs must begin with `/documentation`. Login applications, unrelated external sites, non-English site prefixes, blog/tag/category pages, comparisons and downloadable media are excluded. A public page must also pass the configured relevance filter and robots checks. This scope can be changed in `HOSTS`, `SKIP_PATH`, `TOPICS` and `relevant()` in the source.

## Understand the results before using them

Open the first pages of the document and inspect the coverage counts. Read the Collection issues section. Check `pending_urls` in the JSON if you need to expand the crawl. A missing chapter means no page was assigned to it during this run; it does not prove the feature is unavailable. A single overview page may discuss several capabilities while being assigned to only one chapter.

Vendor claims, discounts, prices, compliance statements and planned features are captured as source statements. They are not verified by the tool. Dates describe when the scraper retrieved a page, not when the vendor last updated it. Canonical URLs are retained as metadata but are not blindly trusted or followed outside the allowed scope. Query strings and fragments are removed during discovery; query-driven content is therefore not collected separately.

If you require a concise narrative with independently checked claims, that is an additional editorial step. The scraper deliberately does not manufacture explanations to make a report appear complete.

## Browser rendering and interactive tours

Auto mode uses the browser when direct HTML extraction contains fewer than 400 characters. Explicit GitBook Markdown pages use a lower 120-character threshold because short documentation index pages can still be useful. A text threshold is only a heuristic: a consent or access page with lots of text may still need manual inspection. Known access challenge phrases are rejected. Use `--mode browser --headed` to inspect rendering. Browser mode waits for DOM content and a meaningful text threshold, then performs four bounded scrolls; it does not depend on `networkidle`.

No consent buttons are clicked by default and banners are not hidden with CSS. If content needs a consent decision, inspect the actual page and pass the selector for the choice you intend to make:

```bash
python aqua_scraper.py --mode browser --headed --max-pages 1 --consent-selector "YOUR_INSPECTED_SELECTOR"
```

Do not paste that placeholder literally. The configured selector applies to each rendered page, so scope such a run to the specific page first. The browser context is temporary and is not saved as a login session.

If a verified tour has a Next control in the main page:

```bash
python aqua_scraper.py --mode browser --seed https://aqua-cloud.io/product-tour/ --max-pages 1 --next-selector "YOUR_INSPECTED_NEXT_SELECTOR" --max-steps 20
```

The loop stops if the control is absent, disabled, unchanged, repeated or reaches the limit. It merges newly observed text blocks. It does not promise to reconstruct slide order or collect visual-only slides. Cross-origin embedded tours are recorded as gaps; automatic iframe traversal and video transcription are not included. Do not use an arbitrary site's Next link without inspecting its action.

Observed browser fetch/XHR URLs, status codes and content types are stored in `network`. Tokens, query strings, cookies, request payloads and API response bodies are not saved. These records are investigation leads, not verified public APIs. The scraper does not replay them or bypass access controls.

## Rebuild a document without crawling again

```bash
python aqua_scraper.py --from-json output/aqua_data.json --output rebuilt
```

For a shorter document while keeping the original collected JSON:

```bash
python aqua_scraper.py --from-json output/aqua_data.json --output concise --max-words-per-page 1200
```

The default word cap is zero, meaning all collected blocks are included. Any configured truncation is labelled. The report writer does not download images, embed screenshots, transcribe videos or extract PDF text. Source tables are rendered as text rows separated by vertical bars, so complex tables should be checked on the linked original page.

## Reliability and limits

- Requests are sequential and paced by host, with a minimum one-second delay by default and any supported robots crawl delay. `--delay 2` increases the interval.
- HTTP 429 and 5xx responses are retried with bounded waits; connection failures have bounded retries. HTTP 401/403 are recorded, not bypassed.
- Redirects are limited, checked against allowed hosts and checked against robots rules. The standard Python robots parser is used; it is not a full Google-style wildcard implementation.
- An unavailable robots file causes that host to be skipped. HTTP 404 is treated as no published robots file. No override switch is supplied.
- Response bodies are limited to 8 MB per page. Browser request metadata is capped at 150 entries per rendered page.
- JSON is checkpointed after every attempted page. Ctrl+C retains a partial result and generates a partial handbook if usable pages exist. Checkpoints support offline document rebuilding, not crawl resumption.
- Exact duplicate page bodies and repeated blocks on the same page are removed. Semantic duplicates across different pages may remain.
- Files in the chosen output folder are overwritten on rerun except the log, which is appended. Use a new output folder for each dated collection.
- Exit code 0 means a handbook was generated, possibly with collection gaps. Exit code 2 means no usable page was collected and no new handbook was generated. Always inspect coverage.

## Troubleshooting

**Browser executable missing:** run `python -m playwright install chromium` with the same interpreter that runs the scraper. HTTP-only mode can still work without Chromium.

**Empty or short result:** inspect `scrape.log`, the JSON issues and robots status. Increase the page/depth budget, run a documentation-focused seed or inspect the page in headed browser mode. Do not lower the content threshold just to make an empty scrape look successful.

**Page shape changed:** adjust `extract_html()` using the actual DOM. The extractor supports headings, paragraphs, lists, tables, summaries and substantial leaf divs. It removes scripts, navigation, forms and known consent chrome. This remains a heuristic, not a guarantee for every future page layout.

**Need a private feature workflow:** this tool does not log in. Use vendor documentation or a separately authorised export from your account.

## Tests

```bash
python -m unittest discover -s tests -v
```

Tests cover URL scope, redirects, robots behaviour, extraction, duplicate removal, source metadata, no-content failure, Word generation and offline rebuilding. See `VALIDATION.md` for the environment-specific checks performed on the delivered version.
