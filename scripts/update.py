"""Daily price-list update for the public site (runs in GitHub Actions).

Usage: python update.py <pricelist.xlsx> <site_dir>
Reads the pricing formula from the PRICING secret (JSON):
  {"url": ..., "div": <divisor>, "add_ils": <fixed add-on in ILS>, "markup": <multiplier>}
selling price (USD) = (list price / div + add_ils / USD_ILS) * markup
USD_ILS = latest Bank of Israel representative rate (falls back to the last one used).
Writes data.json (qty, condition, category, brand, sku, description, price) and meta.json.
Price is shown only for computers; everything else gets null (= contact us).
Neither the list price nor the formula is ever written to the repo.
"""
import csv
import io
import json
import os
import re
import sys
import urllib.request
from datetime import datetime, timedelta, timezone

import openpyxl

def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r'[^\d.\-]', '', str(v))
    try:
        return float(s) if s else None
    except ValueError:
        return None


def condition(section):
    s = (section or '').lower()
    if 'refurb' in s:
        return 'Refurbished'
    if 'open box' in s:
        return 'Open Box'
    if 'used' in s or 'pre-owned' in s:
        return 'Used'
    if 'new' in s:
        return 'New'
    return ''


PROMO = {"Today's Deals": 'deal', 'Price Drop': 'price_drop', 'Just Launched': 'new', 'Just Sold': 'sold'}


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r'[^\d.\-]', '', str(v))
    try:
        return float(s) if s else None
    except ValueError:
        return None


def condition(section):
    s = (section or '').lower()
    if 'refurb' in s:
        return 'Refurbished'
    if 'open box' in s:
        return 'Open Box'
    if 'used' in s or 'pre-owned' in s:
        return 'Used'
    if 'new' in s:
        return 'New'
    return ''


def parse(src, price_fn):
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    offers = {}
    order = []
    stamp = ''
    for ws in wb.worksheets:
        sheet = ws.title.strip()
        promo = PROMO.get(sheet)
        brand = ''
        section = ''
        for row in ws.iter_rows(values_only=True):
            cells = list(row) + [None] * 12
            a, b = cells[0], cells[1]
            filled = [c for c in cells[:12] if c not in (None, '')]
            if not filled:
                continue
            if a not in (None, '') and b in (None, '') and len(filled) == 1:
                txt = str(a).strip()
                m = re.search(r'(\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2})', txt)
                if m and not stamp:
                    stamp = m.group(1)
                if not promo:
                    brand = txt.split('\n')[0].strip()
                continue
            if a in (None, '') and b not in (None, ''):
                section = str(b).strip()
                continue
            price = num(cells[4])
            if a in (None, '') or price is None:
                continue
            sku = str(a).strip()
            desc = str(b or '').strip()
            qty = num(cells[3])
            key = (sku, desc, price, qty, str(cells[7] or '').strip())
            if key not in offers:
                offers[key] = {
                    'sku': sku,
                    'description': desc,
                    'brand': '',
                    'category': '',
                    'section': '',
                    'condition': condition(section) or condition(desc),
                    'location': str(cells[2] or '').strip(),
                    'qty': int(qty) if qty is not None else '',
                                        'sell': price_fn(price),
                    'moq': str(cells[5] or '').strip(),
                    'mxq': str(cells[6] or '').strip(),
                    'warranty': str(cells[7] or '').strip(),
                    'packaging': str(cells[8] or '').strip(),
                    'ship_time': str(cells[9] or '').strip(),
                    'notes': ' '.join(str(c).strip() for c in (cells[10], cells[11]) if c not in (None, '')),
                    'deal': 0, 'price_drop': 0, 'new': 0, 'sold': 0,
                }
                order.append(key)
            o = offers[key]
            if promo:
                o[promo] = 1
                if not o['section']:
                    o['section'] = section
            else:
                if not o['category']:
                    o['category'] = sheet
                    o['brand'] = brand
                    o['section'] = section
                    o['condition'] = condition(section) or o['condition']
    rows = []
    for i, k in enumerate(order, 1):
        o = offers[k]
        if not o['category']:
            o['category'] = 'Promotions only'
        if not o['brand']:
            o['brand'] = (o['description'].split(' ')[0] if o['description'] else '')
        b = o['brand'].strip()
        o['brand'] = b if (len(b) <= 3 or not b.isupper()) else b.title()
        o['id'] = i
        rows.append(o)
    return rows, stamp

MAC = re.compile(r'macbook|imac|mac ?mini|mac studio|mac pro', re.I)


def is_computer(r):
    c, sec, d = r['category'], r.get('section', ''), r['description']
    if c in ('All Laptops', 'Chromebook', 'GPU & PC Components'):
        return True
    if c in ('HP', 'Dell', 'Lenovo'):
        return 'Desktops' in sec
    if c == 'Desktop, LCD & AIO':
        return 'Monitors' not in sec
    if c == 'Apple':
        return bool(MAC.search(d))
    return False


def boi_rate():
    """Latest Bank of Israel representative USD/ILS rate."""
    start = (datetime.now(timezone.utc) - timedelta(days=14)).strftime('%Y-%m-%d')
    url = ('https://edge.boi.gov.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/EXR/1.0/'
           f'RER_USD_ILS?startperiod={start}&format=csv')
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    txt = urllib.request.urlopen(req, timeout=60).read().decode('utf-8-sig')
    rows = list(csv.DictReader(io.StringIO(txt)))
    rows = [r for r in rows if r.get('OBS_VALUE')]
    rows.sort(key=lambda r: r.get('TIME_PERIOD', ''))
    v = float(rows[-1]['OBS_VALUE'])
    return v, rows[-1].get('TIME_PERIOD', '')


def main(xlsx, site):
    cfg = json.loads(os.environ['PRICING'])
    div, add_ils, markup = float(cfg['div']), float(cfg['add_ils']), float(cfg['markup'])
    old = {}
    try:
        old = json.load(open(os.path.join(site, 'meta.json'), encoding='utf-8'))
    except Exception:
        pass
    try:
        rate, rate_date = boi_rate()
        if not 2.5 < rate < 4.5:
            raise ValueError(rate)
    except Exception as e:
        print('rate lookup failed, using last one:', e)
        rate, rate_date = float(old.get('usd_ils') or cfg.get('fallback_rate', 3.25)), old.get('rate_date', '')

    rows, stamp = parse(xlsx, lambda p: round((p / div + add_ils / rate) * markup, 2))
    if len(rows) < 500:
        sys.exit(f'only {len(rows)} rows - file format changed? nothing written')

    out = []
    for r in rows:
        q = r['qty']
        out.append({'qty': int(q) if q not in ('', None) else None, 'condition': r['condition'],
                    'category': r['category'], 'brand': r['brand'], 'sku': r['sku'],
                    'description': r['description'], 'price': r['sell'] if is_computer(r) else None})
    with open(os.path.join(site, 'data.json'), 'w', encoding='utf-8') as f:
        f.write('[\n' + ',\n'.join(json.dumps(x, ensure_ascii=False) for x in out) + '\n]\n')
    updated = ''
    m = re.match(r'(\d{2})-(\d{2})-(\d{4}) (\d{2}:\d{2}:\d{2})', stamp or '')
    if m:
        updated = f'{m.group(3)}-{m.group(1)}-{m.group(2)} {m.group(4)}'
    meta = {'updated': updated or datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S'),
            'count': len(out), 'usd_ils': rate, 'rate_date': rate_date}
    json.dump(meta, open(os.path.join(site, 'meta.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    print(f"rows={len(out)} priced={sum(1 for x in out if x['price'] is not None)} updated={meta['updated']} rate={rate} ({rate_date})")


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
