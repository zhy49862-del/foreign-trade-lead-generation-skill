#!/usr/bin/env python3
"""Clean gosom google-maps-scraper CSV into a deduped customer table.
Usage: python3 clean_results.py <input.csv> <output.csv>
Keeps only business-public fields. Dedupes by website, else phone, else title+address.
"""
import csv, sys, re

def norm(s):
    return re.sub(r'\s+', ' ', (s or '').strip().lower())

def main():
    inp, outp = sys.argv[1], sys.argv[2]
    seen = set()
    rows = []
    with open(inp, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        print('columns:', reader.fieldnames)
        for r in reader:
            title = (r.get('title') or '').strip()
            if not title:
                continue
            website = (r.get('website') or '').strip()
            phone = (r.get('phone') or '').strip()
            address = (r.get('address') or r.get('complete_address') or '').strip()
            key = norm(website) if website else (norm(phone) if phone else norm(title) + '|' + norm(address))
            if key in seen:
                continue
            seen.add(key)
            rows.append({
                '客户名称': title,
                '谷歌类别': (r.get('category') or '').strip(),
                '地址': address,
                '电话': phone,
                '网站': website,
                '评分': r.get('review_rating') or '',
                '评论数': r.get('review_count') or '',
                '谷歌地图链接': r.get('link') or '',
                '搜索词': r.get('input_id') or '',
            })
    rows.sort(key=lambda x: (x['客户名称'].lower()))
    with open(outp, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else ['客户名称'])
        w.writeheader()
        w.writerows(rows)
    print(f'wrote {len(rows)} deduped rows to {outp}')

if __name__ == '__main__':
    main()
