#!/usr/bin/env python3
"""Filter + classify US google-maps-scraper output.
Usage: python3 us_clean.py <input.csv> <output.csv> <stats.txt>
Filter: drop empty title, permanently closed, no phone and no website,
and clearly irrelevant rental (car rental) hits.
Dedupe by website, else phone, else title+address; on duplicates keep
the higher-priority customer type.
"""
import csv, sys, re, collections, json

CITY = {
    'LA': '洛杉矶', 'HOU': '休斯敦', 'DAL': '达拉斯', 'MIA': '迈阿密',
    'ORL': '奥兰多', 'ATL': '亚特兰大', 'CHI': '芝加哥', 'PHX': '凤凰城',
    'DEN': '丹佛', 'SEA': '西雅图', 'Fresno': '弗雷斯诺',
}
def cat_from_input(input_id):
    s = (input_id or '')
    if s.endswith('-mower-dealer') or s.endswith('-garden-machinery'):
        return '经销商/代理商'
    if s.endswith('-agri-machinery'):
        return '农机经销商'
    if s.endswith('-landscaping'):
        return '园林绿化公司'
    if s.endswith('-lawncare'):
        return '庭院养护/草坪服务'
    if s.endswith('-rental'):
        return '机械租赁公司'
    if s.endswith('-gardencenter'):
        return '园艺中心/零售'
    if s.endswith('-construction'):
        return '建筑/工程公司'
    if s.endswith('-tree'):
        return '植被/树木服务'
    return '其他'

PRIORITY = {
    '经销商/代理商': 1, '农机经销商': 2, '机械租赁公司': 3,
    '园林绿化公司': 4, '庭院养护/草坪服务': 5, '植被/树木服务': 6,
    '园艺中心/零售': 7, '建筑/工程公司': 8, '其他': 9,
}
def norm_site(u):
    u = (u or '').strip().lower()
    u = re.sub(r'^https?://', '', u).replace('www.', '', 1).rstrip('/')
    return u
def norm_phone(p):
    return re.sub(r'\D', '', p or '')

def main():
    inp, outp, statp = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(inp, newline='', encoding='utf-8') as f:
        rows = list(csv.DictReader(f))
    raw = len(rows)
    dropped = collections.Counter()
    best = {}
    for r in rows:
        title = (r.get('title') or '').strip()
        if not title:
            dropped['无名称'] += 1; continue
        status = (r.get('status') or '').lower()
        if 'permanently closed' in status or 'permanently_closed' in status:
            dropped['已永久停业'] += 1; continue
        phone = (r.get('phone') or '').strip()
        website = (r.get('website') or '').strip()
        if not phone and not website:
            dropped['无电话无网站'] += 1; continue
        cat_g = (r.get('category') or '').lower()
        if 'car rental' in cat_g or 'car rental agency' in cat_g:
            dropped['汽车租赁（非机械租赁）'] += 1; continue
        if 'restaurant' in cat_g or 'hotel' in cat_g:
            dropped['餐饮酒店等不相关'] += 1; continue
        key = norm_site(website) if website else (norm_phone(phone) if phone else (title.lower() + '|' + (r.get('address') or '').lower()))
        typ = cat_from_input(r.get('input_id'))
        rec = {
            '客户名称': title, '客户类型': typ, '谷歌类别': (r.get('category') or '').strip(),
            '地址': (r.get('address') or '').strip(), '电话': phone, '网站': website,
            '评分': r.get('review_rating') or '', '评论数': r.get('review_count') or '',
            '谷歌地图链接': r.get('link') or '', '来源搜索': r.get('input_id') or '',
        }
        if key not in best or PRIORITY[typ] < PRIORITY[best[key]['客户类型']]:
            if key in best:
                dropped['重复（归入更高优先级类型）'] += 1
            best[key] = rec
        else:
            dropped['重复'] += 1
    out = list(best.values())
    out.sort(key=lambda x: (PRIORITY[x['客户类型']], x['客户名称'].lower()))
    with open(outp, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()) if out else ['客户名称'])
        w.writeheader(); w.writerows(out)
    bycat = collections.Counter(x['客户类型'] for x in out)
    withphone = collections.Counter(x['客户类型'] for x in out if x['电话'])
    withweb = collections.Counter(x['客户类型'] for x in out if x['网站'])
    lines = [f'原始条数：{raw}', f'过滤后条数：{len(out)}', f'过滤掉：{dict(dropped)}', '', '分类统计：']
    for typ, _ in sorted(bycat.items(), key=lambda kv: PRIORITY[kv[0]]):
        lines.append(f"{typ}：{bycat[typ]} 家，其中有电话 {withphone[typ]} 家，有网站 {withweb[typ]} 家")
    lines.append('')
    lines.append(f"有电话合计：{sum(1 for x in out if x['电话'])} 家，有网站合计：{sum(1 for x in out if x['网站'])} 家")
    open(statp, 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print('\n'.join(lines))

if __name__ == '__main__':
    main()
