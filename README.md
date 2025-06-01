# Scraping Tools

This repository contains simple scrapers for Alibaba and Mercari.

## Private Code

The scripts expect a `private_code.py` file to be present in the project
root. This file can contain any functions or data that you prefer not to
share publicly, such as API keys or helper functions.

`private_code.py` is listed in `.gitignore` so it will not be committed.
Create your own version of this file before running the scrapers.

### Required Functions

Core scraping logic has been moved into `private_code.py` to keep it
private. Implement at least the following functions in your own copy of
`private_code.py`:

- `normalize_option_name`, `extract_price_1688`, and `process_excel` for
  the Alibaba scraper.
- `extract_price` and `process_excel` for the Mercari scraper.
