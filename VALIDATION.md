# Validation

Checked on 31 October 2026 with Python 3.12 on Linux.

## Passed

- 17 automated tests passed.
- Live HTTP HTML extraction succeeded for the six product pages listed below.
- Live HTTP Markdown extraction succeeded for the documentation root and the short General page.
- Word generation and offline rebuilding were exercised.
- The five-page user guide was rendered and every page visually reviewed.

## Live sources

- https://aqua-cloud.io/product-tour/ — http HTML; 6005 extracted characters
- https://docs.aqua-cloud.io/documentation — http Markdown; 1896 extracted characters
- https://aqua-cloud.io/requirements-engineering-tool/ — http HTML; 7865 extracted characters
- https://aqua-cloud.io/test-case-management-tool/ — http HTML; 8559 extracted characters
- https://aqua-cloud.io/bug-tracking-tools/ — http HTML; 8643 extracted characters
- https://aqua-cloud.io/test-automation-tools/ — http HTML; 6931 extracted characters
- https://aqua-cloud.io/pricing/ — http HTML; 6795 extracted characters
- https://docs.aqua-cloud.io/documentation/web-client/general.md — http Markdown; 323 extracted characters

The broader eight-URL run initially skipped the 323-character General documentation page as too thin. After adding the Markdown-specific threshold and its regression test, a focused live rerun collected that page and the product tour with zero issues. These checks do not establish completeness of a 50-page or 100-page crawl.

## Not verified in this environment

- The Playwright package installed, but Chromium downloads returned invalid archives. Live browser rendering, consent interaction and Next-button traversal remain unverified.
- Windows and macOS launchers are provided but were not executed on those operating systems.
- Private product screens, iframe tours, image OCR, PDF extraction and video transcription are outside this implementation.

## Installed package versions

- lxml: 6.1.1
- python-docx: 1.2.0
- playwright: 1.64.1
- markdown-it-py: 4.2.0
