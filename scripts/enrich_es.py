#!/usr/bin/env python3
"""Enrich ES leads: fetch each company website, extract emails,
contact person, social links, and classify by website content (English).
Resumable via JSONL checkpoint. --retry-failed reprocesses unreachable rows.
Usage: ENRICH_WORKERS=4 ENRICH_TIMEOUT=12 python3 enrich_uk.py <input_csv> <checkpoint_jsonl> <output_csv> [--retry-failed]
"""
import csv, json, os, re, sys, time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import requests

UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36'}
EMAIL_RE = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}')
BAD_EMAIL_LOCAL = {'noreply', 'no-reply', 'abuse', 'postmaster', 'webmaster', 'donotreply', 'newsletter'}
WORKERS = int(os.environ.get('ENRICH_WORKERS', '4'))
TIMEOUT = int(os.environ.get('ENRICH_TIMEOUT', '12'))

NAME_STOP = {'Ltd', 'Limited', 'Company', 'Group', 'Contact', 'Email', 'Phone', 'Address',
             'VAT', 'Registered', 'Office', 'Sales', 'Service', 'Team', 'Our', 'The', 'And'}
NAME_PATTERNS = [
    re.compile(r'(?:Administrador|Titular|Representante legal)\s*[:\-–]?\s*(?:Sr\.?|Sra\.?|D\.?|Dña\.?)?\s*([A-Z][a-záéíóúñü]+(?:\s+[A-Z][a-záéíóúñü]+){0,2})'),
    re.compile(r'(?:Director|Directora|Presidente|Gerente|Responsable)\s*[:\-–]?\s*(?:Sr\.?|Sra\.?|D\.?|Dña\.?)?\s*([A-Z][a-záéíóúñü]+(?:\s+[A-Z][a-záéíóúñü]+){0,2})'),
    re.compile(r'Fundador\s*[:\-–]?\s*(?:Sr\.?|Sra\.?|D\.?|Dña\.?)?\s*([A-Z][a-záéíóúñü]+(?:\s+[A-Z][a-záéíóúñü]+){0,2})'),
]

KEYWORDS = {
    'mower': ['cortacésped', 'cortacesped', 'robot cortacésped', 'maquinaria de jardín', 'jardinería'],
    'excavator': ['miniexcavadora', 'mini excavadora', 'excavadora', 'retroexcavadora'],
    'forklift': ['carretilla elevadora', 'carretilla', 'transpaleta', 'toro'],
    'loader': ['pala cargadora', 'pala', 'cargadora'],
    'machinery_dealer': ['concesionario', 'maquinaria de construcción', 'maquinaria de obra', 'venta de maquinaria', 'distribuidor'],
    'rental': ['alquiler de maquinaria', 'alquiler de máquinas', 'alquiler', 'alquilar'],
    'landscaping': ['jardinero', 'mantenimiento de jardines', 'diseño de jardines', 'zonas verdes'],
    'construction': ['construcción', 'constructora', 'empresa constructora', 'movimiento de tierras', 'obras públicas'],
    'agriculture': ['agrícola', 'maquinaria agrícola', 'tractor'],
    'irrelevant': ['supermercado', 'restaurante', 'hotel', 'ropa', 'panadería', 'farmacia', 'muebles', 'electrodomésticos', 'peluquería'],
}

SOCIAL = {
    'Facebook': re.compile(r'https?://(?:www\.)?facebook\.com/[A-Za-z0-9._\-/]+', re.I),
    'Instagram': re.compile(r'https?://(?:www\.)?instagram\.com/[A-Za-z0-9._\-/]+', re.I),
    'LinkedIn': re.compile(r'https?://(?:www\.)?linkedin\.com/(?:company|in)/[A-Za-z0-9._\-/]+', re.I),
    'YouTube': re.compile(r'https?://(?:www\.)?youtube\.com/(?:channel|user|c)/[A-Za-z0-9._\-/]+', re.I),
    'TikTok': re.compile(r'https?://(?:www\.)?tiktok\.com/@[A-Za-z0-9._\-/]+', re.I),
    'X/Twitter': re.compile(r'https?://(?:www\.)?(?:x\.com|twitter\.com)/[A-Za-z0-9_]+', re.I),
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

def fetch(url, timeout=None):
    timeout = timeout or TIMEOUT
    for attempt in range(2):
        try:
            r = requests.get(url, headers=UA, timeout=timeout, allow_redirects=True)
            if r.status_code == 200 and 'text/html' in r.headers.get('Content-Type', ''):
                return r.text
            if r.status_code in (403, 429, 503):
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
                or re.fullmatch(r'[0-9a-f]{32}', local)):
            continue
        if dom.lower() in ('example.com', 'example.co.uk', 'domain.co.uk'):
            continue
        found.add(m.lower())
    def rank(e):
        dom = e.rsplit('@', 1)[1]
        score = 0
        if domain and dom == domain:
            score += 10
        if e.startswith(('info@', 'contacto@', 'comercial@', 'ventas@', 'administracion@', 'hola@')):
            score += 5
        return -score
    return sorted(found, key=rank)[:5]

def extract_person(html):
    text = strip_tags(html)
    for pat in NAME_PATTERNS:
        for m in pat.finditer(text):
            name = m.group(1).strip()
            parts = name.split()
            if len(parts) < 2 or any(p in NAME_STOP for p in parts):
                continue
            if len(name) < 40:
                return name
    return ''

def extract_social(html):
    out = {}
    for key, pat in SOCIAL.items():
        m = pat.search(html)
        out[key] = m.group(0).rstrip('/') if m else ''
    return out

def find_subpages(html, base):
    out = []
    for href in re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.I):
        low = href.lower()
        if 'contact' in low or 'contacto' in low:
            out.append(('contact', urljoin(base, href)))
        elif 'about' in low or 'quienes-somos' in low or 'empresa' in low:
            out.append(('about', urljoin(base, href)))
        elif 'team' in low or 'equipo' in low:
            out.append(('team', urljoin(base, href)))
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

def suggest_grade(hits):
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

GOOGLE_CAT_F = ['supermarché', 'restaurant', 'hôtel', 'vêtements', 'boulangerie', 'pharmacie', 'magasin de meubles', 'grand magasin', 'épicerie', 'magasin']
GOOGLE_CAT_C = ['jardinerie', 'pépinière', 'paysagiste', 'entretien de jardin']
GOOGLE_CAT_A = ['quincaillerie', 'magasin de bricolage', 'fournisseur de matériel agricole', 'machines']

def phone_type(phone):
    digits = re.sub(r'\D', '', phone or '')
    if digits.startswith('34'):
        digits = digits[2:]
    if (digits.startswith('6') or digits.startswith('7')) and len(digits) == 9:
        return '手机', digits
    if digits:
        return '座机', digits
    return '', ''

def enrich_row(row):
    site = clean_url(row.get('网站', ''))
    result = {'全部邮箱': '', '主邮箱': '', '联系人': '', '网站分类': '', '建议分级': '', '删除原因': '',
              '网站可达': '否', 'Facebook': '', 'Instagram': '', 'LinkedIn': '', 'YouTube': '', 'TikTok': '', 'X/Twitter': '',
              '号码类型': '', 'WhatsApp': '', 'WhatsApp校验': ''}
    ptype, digits = phone_type(row.get('电话', ''))
    result['号码类型'] = ptype
    if ptype == '手机':
        result['WhatsApp'] = 'https://wa.me/34' + digits
        result['WhatsApp校验'] = '未验证（仅号码判断，需在WhatsApp里确认是否注册）'
    elif ptype == '座机':
        result['WhatsApp校验'] = '座机号，不适用'
    if not site:
        result['网站可达'] = '无网站'
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
    result['全部邮箱'] = '; '.join(emails)
    result['主邮箱'] = emails[0] if emails else ''
    social = extract_social(all_html)
    result.update(social)
    text_low = strip_tags(all_html).lower()
    hits = classify(text_low)
    cats = [t for t, w in hits.items() if w]
    result['网站分类'] = ','.join(cats)
    result['联系人'] = extract_person(all_html)
    g, why = suggest_grade(hits)
    if not g:
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

EXTRA_COLS = ['全部邮箱', '主邮箱', '联系人', '网站分类', '建议分级', '删除原因', '网站可达',
              'Facebook', 'Instagram', 'LinkedIn', 'YouTube', 'TikTok', 'X/Twitter', '号码类型', 'WhatsApp', 'WhatsApp校验']

def main():
    inp = sys.argv[1]
    ckpt = sys.argv[2]
    outp = sys.argv[3]
    retry_failed = '--retry-failed' in sys.argv
    os.makedirs(os.path.dirname(ckpt), exist_ok=True)
    rows = list(csv.DictReader(open(inp, encoding='utf-8-sig')))
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
                    res = {c: '' for c in EXTRA_COLS}
                    res['网站可达'] = '出错:' + str(e)[:60]
                done[key] = {'_key': key, **res}
                cf.write(json.dumps({'_key': key, **res}, ensure_ascii=False) + '\n')
                cf.flush()
                n += 1
                if n % 100 == 0:
                    print(f'{n}/{len(todo)} elapsed {time.time()-t0:.0f}s', flush=True)
    fields = list(rows[0].keys()) + EXTRA_COLS
    with open(outp, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            key = (r.get('客户名称', '') + '|' + r.get('地址', ''))[:200]
            d = done.get(key, {})
            out = dict(r)
            for col in EXTRA_COLS:
                out[col] = d.get(col, '')
            w.writerow(out)
    import collections
    print('grades:', dict(collections.Counter(done.get((r.get('客户名称','')+'|'+r.get('地址',''))[:200], {}).get('建议分级','') for r in rows)))
    print('with email:', sum(1 for r in rows if done.get((r.get('客户名称','')+'|'+r.get('地址',''))[:200], {}).get('全部邮箱')))
    print('with person:', sum(1 for r in rows if done.get((r.get('客户名称','')+'|'+r.get('地址',''))[:200], {}).get('联系人')))
    print('mobile:', sum(1 for r in rows if done.get((r.get('客户名称','')+'|'+r.get('地址',''))[:200], {}).get('号码类型')=='手机'))
    print('done, wrote', outp, flush=True)

if __name__ == '__main__':
    main()
