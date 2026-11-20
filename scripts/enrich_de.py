#!/usr/bin/env python3
"""Enrich German leads: fetch each company website, extract emails,
contact person (Impressum), and classify by website content.
Resumable: results appended to JSONL checkpoint; re-run skips done rows.
Usage: python3 enrich_de.py [input_csv] [checkpoint_jsonl] [output_csv]
"""
import csv, json, os, re, sys, time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import requests

UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36'}
EMAIL_RE = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}')
BAD_EMAIL_LOCAL = {'noreply', 'no-reply', 'abuse', 'postmaster', 'webmaster', 'donotreply', 'newsletter'}
NAME_STOP = {'Umsatzsteuer', 'Impressum', 'Handelsregister', 'Amtsgericht', 'Telefon', 'Telefax',
             'Straße', 'Strasse', 'GmbH', 'AG', 'KG', 'UG', 'Im', 'In', 'Der', 'Die', 'Das', 'Und',
             'E-Mail', 'Email', 'Adresse', 'Sitz', 'Registergericht', 'USt-IdNr', 'Vertreten'}
NAME_PATTERNS = [
    re.compile(r'Gesch[äa]ftsf[üu]hrer(?:in)?\s*[:\-–]?\s*(?:Dr\.\s*)?([A-ZÄÖÜ][a-zäöüß]+)(?:\s+(von|van|zu))?\s+([A-ZÄÖÜ][a-zäöüß]+)'),
    re.compile(r'Gesch[äa]ftsf[üu]hrung\s*[:\-–]?\s*(?:Dr\.\s*)?([A-ZÄÖÜ][a-zäöüß]+)(?:\s+(von|van|zu))?\s+([A-ZÄÖÜ][a-zäöüß]+)'),
    re.compile(r'Inhaber(?:in)?\s*[:\-–]?\s*(?:Dr\.\s*)?([A-ZÄÖÜ][a-zäöüß]+)(?:\s+(von|van|zu))?\s+([A-ZÄÖÜ][a-zäöüß]+)'),
    re.compile(r'Ansprechpartner(?:in)?\s*[:\-–]?\s*(?:Dr\.\s*)?([A-ZÄÖÜ][a-zäöüß]+)(?:\s+(von|van|zu))?\s+([A-ZÄÖÜ][a-zäöüß]+)'),
]

# fallback classification from Google category when the website gives no keywords
GOOGLE_CAT_F = ['discount store', 'supermarket', 'grocery', 'restaurant', 'hotel', 'clothing', 'bakery', 'pharmacy', 'florist', 'furniture store', 'department store', 'convenience store', 'drugstore']
GOOGLE_CAT_C = ['garden center', 'nursery', 'landscaper', 'lawn care service']
GOOGLE_CAT_A = ['hardware store', 'home goods store', 'tool store', 'farm equipment supplier', 'machinery']

# keyword -> (category tag, weight)
KEYWORDS = {
    'mower': ['rasenmäher', 'rasenmaeher', 'mähroboter', 'rasentraktor', 'rasenmäher-händler', 'gartengeräte'],
    'excavator': ['minibagger', 'bagger', 'kettenbagger', 'mobilbagger'],
    'forklift': ['gabelstapler', 'stapler', 'hubwagen', 'flurförderzeug'],
    'loader': ['radlader', 'hoflader', 'kompaktlader', 'teleskoplader'],
    'machinery_dealer': ['baumaschinen', 'maschinenhandel', 'maschinenhändler', 'landmaschinen', 'motorgeräte'],
    'rental': ['verleih', 'vermietung', 'mieten', 'mietpark', 'arbeitsbühnen'],
    'landscaping': ['gartenbau', 'landschaftsbau', 'galabau', 'garten- und landschaftsbau', 'gärtnerei', 'landschaftspflege'],
    'construction': ['bauunternehmen', 'hochbau', 'tiefbau', 'straßenbau', 'erdarbeiten', 'abbruch'],
    'agriculture': ['landtechnik', 'agrar', 'landwirtschaft'],
    'irrelevant': ['supermarkt', 'discounter', 'restaurant', 'hotel', 'bekleidung', 'modehaus', 'lebensmittel', 'bäckerei', 'friseur', 'apotheke'],
}

def clean_url(u):
    if not u:
        return ''
    u = u.strip()
    if u.startswith('//'):
        u = 'https:' + u
    if not u.startswith('http'):
        u = 'https://' + u
    return u.split('?utm_')[0].split('&utm_')[0]

WORKERS = int(os.environ.get('ENRICH_WORKERS', '6'))
TIMEOUT = int(os.environ.get('ENRICH_TIMEOUT', '12'))

def fetch(url, timeout=None):
    timeout = timeout or TIMEOUT
    for attempt in range(2):
        try:
            r = requests.get(url, headers=UA, timeout=timeout, allow_redirects=True)
            if r.status_code == 200 and 'text/html' in r.headers.get('Content-Type', ''):
                return r.text
            if r.status_code in (403, 429, 503):
                # bot protection: retrying immediately is useless
                return ''
        except Exception:
            time.sleep(2)
    if url.startswith('https://'):
        try:
            r = requests.get('http://' + url[len('https://'):], headers=UA, timeout=timeout, allow_redirects=True)
            if r.status_code == 200 and 'text/html' in r.headers.get('Content-Type', ''):
                return r.text
        except Exception:
            pass
    return ''

def strip_tags(html):
    html = re.sub(r'<(script|style)[^>]*>.*?</\1>', ' ', html, flags=re.S | re.I)
    html = re.sub(r'<[^>]+>', ' ', html)
    return re.sub(r'\s+', ' ', html)

def extract_emails(html, domain):
    found = set()
    for m in EMAIL_RE.findall(html):
        local, dom = m.rsplit('@', 1)
        if local.lower() in BAD_EMAIL_LOCAL:
            continue
        # Junk the regex also matches: Sentry error-tracking IDs and
        # image/file names in page HTML (e.g. logo@2x.png, ...@sentry.io).
        if ('sentry' in dom.lower()
                or dom.lower().rsplit('.', 1)[-1] in ('png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'js', 'css', 'map')
                or dom.lower().endswith('.de.de')
                or re.fullmatch(r'[0-9a-f]{32}', local)):
            continue
        if dom.lower() in ('example.com', 'example.de', 'domain.de'):
            continue
        found.add(m.lower())
    # prefer same-domain emails, then generic contact ones
    def rank(e):
        dom = e.rsplit('@', 1)[1]
        score = 0
        if domain and dom == domain:
            score += 10
        if e.startswith(('info@', 'kontakt@', 'anfrage@', 'sales@', 'vertrieb@', 'service@')):
            score += 5
        return -score
    return sorted(found, key=rank)[:5]

def extract_person(html):
    text = strip_tags(html)
    for pat in NAME_PATTERNS:
        for m in pat.finditer(text):
            first = m.group(1)
            mid = m.group(2) or ''
            last = m.group(3)
            if first in NAME_STOP or last in NAME_STOP:
                continue
            if len(first) < 2 or len(last) < 2:
                continue
            name = (first + ' ' + (mid + ' ' if mid else '') + last).strip()
            if len(name) < 40:
                return name
    return ''

def find_subpages(html, base):
    out = []
    for href in re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.I):
        low = href.lower()
        if 'impressum' in low:
            out.append(('impressum', urljoin(base, href)))
        elif 'kontakt' in low:
            out.append(('kontakt', urljoin(base, href)))
    seen = set(); uniq = []
    for kind, u in out:
        if u not in seen:
            seen.add(u); uniq.append((kind, u))
    return uniq[:3]

def classify(text_low):
    hits = {}
    for tag, words in KEYWORDS.items():
        hits[tag] = [w for w in words if w in text_low]
    return hits

def suggest_grade(row_type, hits):
    """Draft A-F grade (pending user confirmation of scheme).
    Irrelevant words (clothing/food/hotel in nav/footer) only count as F
    when NO machinery-related keyword appears anywhere on the site."""
    relevant = (hits['machinery_dealer'] or hits['excavator'] or hits['forklift']
                or hits['loader'] or hits['mower'] or hits['rental']
                or hits['landscaping'] or hits['construction'] or hits['agriculture'])
    if hits['irrelevant'] and not relevant:
        return 'F', '网站内容与机械无关：' + ','.join(hits['irrelevant'][:3])
    dealer_words = hits['machinery_dealer'] or hits['excavator'] or hits['forklift'] or hits['loader'] or hits['mower']
    if dealer_words:
        return 'A', '经销商，涉及：' + ','.join(sorted(set(hits['machinery_dealer'] + hits['excavator'] + hits['forklift'] + hits['loader'] + hits['mower']))[:4])
    if hits['rental']:
        return 'B', '租赁：' + ','.join(hits['rental'][:3])
    if hits['landscaping']:
        return 'C', '园林绿化：' + ','.join(hits['landscaping'][:3])
    if hits['construction']:
        return 'D', '建筑工程：' + ','.join(hits['construction'][:3])
    if hits['agriculture']:
        return 'E', '农机：' + ','.join(hits['agriculture'][:3])
    return '', '网站未读取到分类关键词'

def enrich_row(row):
    site = clean_url(row.get('网站', ''))
    result = {'邮箱': '', '联系人': '', '网站分类': '', '建议分级': '', '删除原因': '', '网站可达': '否'}
    if not site:
        result['网站可达'] = '无网站'
        g, why = suggest_grade(row.get('客户类型', ''), {k: [] for k in KEYWORDS})
        result['建议分级'] = g
        return result
    domain = urlparse(site).netloc.lower().replace('www.', '')
    home = fetch(site)
    if not home:
        result['网站可达'] = '打不开'
        return result
    result['网站可达'] = '是'
    all_html = home
    for kind, url in find_subpages(home, site):
        sub = fetch(url)
        if sub:
            all_html += '\n' + sub
    emails = extract_emails(all_html, domain)
    result['邮箱'] = '; '.join(emails)
    text_low = strip_tags(all_html).lower()
    hits = classify(text_low)
    cats = [t for t, w in hits.items() if w]
    result['网站分类'] = ','.join(cats)
    person = extract_person(all_html)
    result['联系人'] = person
    g, why = suggest_grade(row.get('客户类型', ''), hits)
    if not g:
        # fallback: Google category decides when website text gives nothing
        gcat = (row.get('谷歌类别', '') or '').lower()
        if any(w in gcat for w in GOOGLE_CAT_F):
            g, why = 'F', '谷歌类别与机械无关：' + row.get('谷歌类别', '')
        elif any(w in gcat for w in GOOGLE_CAT_C):
            g, why = 'C', '谷歌类别：' + row.get('谷歌类别', '')
        elif any(w in gcat for w in GOOGLE_CAT_A):
            g, why = 'A', '谷歌类别：' + row.get('谷歌类别', '')
    result['建议分级'] = g
    if g == 'F':
        result['删除原因'] = why
    return result

def main():
    inp = sys.argv[1] if len(sys.argv) > 1 else 'out/by_country/DE_德国.csv'
    ckpt = sys.argv[2] if len(sys.argv) > 2 else 'out/enrich/DE_enriched.jsonl'
    outp = sys.argv[3] if len(sys.argv) > 3 else 'out/enrich/DE_enriched.csv'
    os.makedirs(os.path.dirname(ckpt), exist_ok=True)
    rows = list(csv.DictReader(open(inp, encoding='utf-8-sig')))
    retry_failed = '--retry-failed' in sys.argv
    done = {}
    if os.path.exists(ckpt):
        with open(ckpt, encoding='utf-8') as f:
            for line in f:
                try:
                    d = json.loads(line)
                    done[d['_key']] = d
                except Exception:
                    pass
    if retry_failed:
        n_fail = sum(1 for d in done.values() if d.get('网站可达') == '打不开')
        done = {k: v for k, v in done.items() if v.get('网站可达') != '打不开'}
        # rewrite checkpoint without the failed entries
        with open(ckpt, 'w', encoding='utf-8') as f:
            for v in done.values():
                f.write(json.dumps(v, ensure_ascii=False) + '\n')
        print(f'retry-failed: dropped {n_fail} unreachable entries', flush=True)
    print(f'total {len(rows)}, already done {len(done)}', flush=True)
    todo = []
    for i, r in enumerate(rows):
        key = (r.get('客户名称', '') + '|' + r.get('地址', ''))[:200]
        if key not in done:
            todo.append((i, key, r))
    print(f'to process {len(todo)}', flush=True)
    t0 = time.time()
    with open(ckpt, 'a', encoding='utf-8') as cf:
        with ThreadPoolExecutor(max_workers=WORKERS) as ex:
            futs = {ex.submit(enrich_row, r): (i, key) for i, key, r in todo}
            n = 0
            for fut in as_completed(futs):
                i, key = futs[fut]
                try:
                    res = fut.result()
                except Exception as e:
                    res = {'邮箱': '', '联系人': '', '网站分类': '', '建议分级': '', '删除原因': '', '网站可达': '出错:' + str(e)[:60]}
                done[key] = {'_key': key, **res}
                cf.write(json.dumps({'_key': key, **res}, ensure_ascii=False) + '\n')
                cf.flush()
                n += 1
                if n % 100 == 0:
                    print(f'{n}/{len(todo)} elapsed {time.time()-t0:.0f}s', flush=True)
    fields = list(rows[0].keys()) + ['邮箱', '联系人', '网站分类', '建议分级', '删除原因', '网站可达']
    with open(outp, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            key = (r.get('客户名称', '') + '|' + r.get('地址', ''))[:200]
            d = done.get(key, {})
            out = dict(r)
            for col in ['邮箱', '联系人', '网站分类', '建议分级', '删除原因', '网站可达']:
                out[col] = d.get(col, '')
            w.writerow(out)
    import collections
    print('grades:', dict(collections.Counter(done.get((r.get('客户名称','')+'|'+r.get('地址',''))[:200], {}).get('建议分级','') for r in rows)))
    print('with email:', sum(1 for r in rows if done.get((r.get('客户名称','')+'|'+r.get('地址',''))[:200], {}).get('邮箱')))
    print('with person:', sum(1 for r in rows if done.get((r.get('客户名称','')+'|'+r.get('地址',''))[:200], {}).get('联系人')))
    print('done, wrote', outp, flush=True)

if __name__ == '__main__':
    main()
