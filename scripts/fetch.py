"""Download the supplier price list (.xlsx) with a headless browser.

The source page starts the download by itself when it is opened.
Usage: python fetch.py <output.xlsx>
The page URL comes from the PRICING secret (JSON, key "url"), never from the repo.
"""
import json
import os
import sys

from playwright.sync_api import sync_playwright


def main(out):
    cfg = json.loads(os.environ['PRICING'])
    url = cfg['url']
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(accept_downloads=True, viewport={'width': 1400, 'height': 900},
                                  user_agent=('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                                              '(KHTML, like Gecko) Chrome/131.0 Safari/537.36'))
        page = ctx.new_page()
        with page.expect_download(timeout=180_000) as dl:
            page.goto(url, wait_until='domcontentloaded', timeout=120_000)
        d = dl.value
        d.save_as(out)
        print('downloaded', os.path.getsize(out), 'bytes')
        browser.close()


if __name__ == '__main__':
    main(sys.argv[1])
