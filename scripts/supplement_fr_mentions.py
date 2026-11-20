#!/usr/bin/env python3
"""Supplementary pass for France: French company sites must publish a
'mentions légales' page, which often names the gérant (manager) and a
contact email. The main enrich_fr pass only read homepage + contact /
about / team pages, so re-fetch the homepage, follow the mentions-légales
link, and extract any extra emails + person names.

Scope: rows with 网站可达 == '是' and (website grade A/B, or reachable
but ungraded). Grades are NOT touched; only 全部邮箱 / 主邮箱 / 联系人
are supplemented. Resumable via JSONL checkpoint.
Usage: python3 supplement_fr_mentions.py <enriched_csv> <checkpoint_jsonl> <output_csv>
"""
import csv, json, os, re, sys, time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from enrich_fr import fetch, extract_emails, clean_url, strip_tags, NAME_PATTERNS
from email_junk import clean_email_list

PUB_PATTERN = re.compile(
    r"Directeur de la publication\s*[:\-–]?\s*(?:M\.?|Mme\.?|Mlle\.?)?\s*"
    r"([A-Z][a-zàâäéèêëîïôöùûüç]+(?:\s+[A-Z][a-zàâäéèêëîïôöùûüç]+){0,2})")
GERANT_PATTERN = re.compile(
    r"(?:repr[eé]sent[eé]e? par|G[eé]rant)\s*[:\-–]?\s*(?:M\.?|Mme\.?|Mlle\.?)?\s*"
    r"([A-Z][a-zàâäéèêëîïôöùûüç]+(?:\s+[A-Z][a-zàâäéèêëîïôöùûüç]+){0,2})")

NAME_STOP = {'Ltd', 'Limited', 'Company', 'Group', 'Contact', 'Email', 'Phone',
             'SARL', 'SAS', 'EURL', 'SA', 'Sarl', 'Sas'}


def extract_person_fr(html):
    text = strip_tags(html)
    for pat in [PUB_PATTERN, GERANT_PATTERN] + NAME_PATTERNS:
        for m in pat.finditer(text):
            name = m.group(1).strip()
            parts = name.split()
            if len(parts) < 2 or any(p in NAME_STOP for p in parts):
                continue
            if len(name) < 40:
                return name
    return ''


def find_mentions(html, base):
    for href in re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.I):
        low = href.lower()
        if 'mention' in low or 'legal' in low:
            return urljoin(base, href)
    return ''


def process(row):
    site = clean_url(row.get('网站', ''))
    out = {'new_emails': '', 'new_person': ''}
    if not site:
        return out
    home = fetch(site)
    if not home:
        return out
    url = find_mentions(home, site)
    if not url:
        return out
    page = fetch(url)
    if not page:
        return out
    from urllib.parse import urlparse
    domain = urlparse(site).netloc.lower().replace('www.', '')
    emails = extract_emails(page, domain)
    out['new_emails'] = clean_email_list('; '.join(emails))
    out['new_person'] = extract_person_fr(page)
    return out


def main():
    inp, ckpt, outp = sys.argv[1], sys.argv[2], sys.argv[3]
    rows = list(csv.DictReader(open(inp, encoding='utf-8-sig')))
    done = {}
    if os.path.exists(ckpt):
        for line in open(ckpt, encoding='utf-8'):
            try:
                d = json.loads(line)
                done[d['_key']] = d
            except Exception:
                pass
    todo = []
    for r in rows:
        if r.get('网站可达') != '是':
            continue
        if r.get('建议分级') not in ('A', 'B', ''):
            continue
        key = (r.get('客户名称', '') + '|' + r.get('地址', ''))[:200]
        if key not in done:
            todo.append((key, r))
    print(f'total {len(rows)}, target {len(todo) + len(done)}, already done {len(done)}, to process {len(todo)}', flush=True)
    t0 = time.time()
    with open(ckpt, 'a', encoding='utf-8') as cf:
        with ThreadPoolExecutor(max_workers=4) as ex:
            futs = {ex.submit(process, r): key for key, r in todo}
            n = 0
            for fut in as_completed(futs):
                key = futs[fut]
                try:
                    res = fut.result()
                except Exception as e:
                    res = {'new_emails': '', 'new_person': ''}
                done[key] = {'_key': key, **res}
                cf.write(json.dumps({'_key': key, **res}, ensure_ascii=False) + '\n')
                cf.flush()
                n += 1
                if n % 100 == 0:
                    print(f'{n}/{len(todo)} elapsed {time.time()-t0:.0f}s', flush=True)
    n_email = n_person = 0
    with open(outp, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        for r in rows:
            key = (r.get('客户名称', '') + '|' + r.get('地址', ''))[:200]
            d = done.get(key)
            if d:
                old = [p.strip() for p in (r.get('全部邮箱') or '').split(';') if p.strip()]
                new = [p.strip() for p in (d.get('new_emails') or '').split(';') if p.strip()]
                merged = old + [e for e in new if e.lower() not in [o.lower() for o in old]]
                if merged != old:
                    n_email += 1
                    r['全部邮箱'] = '; '.join(merged)
                    r['主邮箱'] = merged[0] if merged else ''
                if not (r.get('联系人') or '').strip() and d.get('new_person'):
                    r['联系人'] = d['new_person']
                    n_person += 1
            w.writerow(r)
    print(f'rows with new emails: {n_email}, rows with new person: {n_person}')
    print('done, wrote', outp, flush=True)


if __name__ == '__main__':
    main()
