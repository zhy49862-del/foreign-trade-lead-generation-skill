#!/usr/bin/env python3
"""Build IT XLSX in the same 4-tier format as the Germany/UK/FR tables.

Input:  workspace/google-maps-scraper-run/out/enrich/IT_enriched.csv
        (2092 rows; enrichment finished 2026-11-08 ~01:39 UTC)
Steps, in order:
  1. Clean fake emails with email_junk.clean_email_list (Sentry IDs,
     image filenames like logo@2x.png). Recompute 主邮箱 from cleaned list.
  2. Clean contact-person names: strip trailing title/place words
     (Director/General Manager/Operations/Depot/...) and drop fragments
     (single names, 'About Solly', 'Five Stars', 'View Profile', ...).
     Names are extracted from website text, NOT verified - column header
     says 待核对.
  3. Final grade = website grade (建议分级) when present, else fallback
     by 客户类型 / 谷歌分类 (same A-F scheme as Germany; fallback is
     labelled in 判断来源, never presented as website-read).
  4. Tier mapping to the image's 4 tiers, as in Germany:
     A=A, B=B+E, C=C+D, D=F.
Output: HENBO意大利客户审核_2092家.xlsx (12 sheets) + IT_enriched_cleaned.csv
"""
import csv, re, sys, collections
from pathlib import Path
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path.home() / "workspace/google-maps-scraper-run"))
from email_junk import clean_email_list, is_junk_email

SRC = Path.home() / "workspace/google-maps-scraper-run/out/enrich/IT_enriched.csv"
OUTDIR = Path.home() / "workspace/goals/rc-mower-export-lead-prospecting/files"
OUT = OUTDIR / "HENBO意大利客户审核_2092家.xlsx"
CLEANED = OUTDIR / "IT_enriched_cleaned.csv"

# ---------- person cleaning ----------
TRAILING = {'director', 'general', 'manager', 'financial', 'operations', 'depot',
            'project', 'services', 'marketing', 'commercial', 'senior', 'estimator',
            'head', 'lawn', 'colchester', 'central', 'land', 'managing',
            'directeur', 'général', 'gérant', 'président', 'responsable',
            'téléphone', 'tél', 'adjoint', 'chariots', 'matériels', 'vente',
            'sud', 'nord', 'est', 'ouest', 'magasin', 'technique', 'administrative',
            'amministratore', 'titolare', 'direttore', 'presidente',
            'delegato', 'tecnico', 'commerciale', 'generale', 'consulente',
            'capo', 'logistico'}
LEADING = {'administrative', 'technique', 'commercial', 'marketing', 'magasin',
           'directeur', 'général', 'general', 'responsable', 'gérant', 'président',
           'accueil', 'vente', 'espace', 'service', 'directrice',
           'madame', 'monsieur', 'mme', 'm.', 'mr', 'mlle',
           'delegato', 'tecnico', 'commerciale', 'generale', 'consulente',
           'amministratore', 'titolare', 'direttore', 'presidente', 'dott',
           'sig', 'sig.ra', 'ing', 'geom', 'rag'}
GENERIC = {'about', 'read', 'view', 'five', 'stars', 'technical', 'motor', 'chief',
           'executive', 'officer', 'pallet', 'force', 'test', 'tube', 'products',
           'favourite', 'thing', 'severn', 'insulation', 'menu', 'as', 'north',
           'meet', 'close', 'send', 'all', 'bio', 'profile', 'testimonials',
           'scotland', 'us', 'est', 'jnr', 'snr', 'ltd', 'limited', 'company',
           'group', 'contact', 'email', 'team', 'office', 'sales', 'finance',
           'général', 'general', 'commercial', 'marketing', 'technique',
           'administrative', 'magasin', 'téléphone', 'tél', 'espace', 'vente',
           'matériels', 'chariots', 'accueil', 'équipe', 'une', 'des', 'les',
           'du', 'de', 'la', 'le', 'aix', 'sarl', 'sas', 'eurl',
           'madame', 'monsieur', 'mme', 'hébergeur', 'hebergeur', 'webmaster',
           'mr', 'm', 'mlle', 'ms', 'dr', 'community', 'réalisation', 'realisation',
           'assistenza', 'tecnica', 'tecnico', 'consenso', 'newsletter', 'acconsento',
           'prevenzione', 'corruzione', 'post', 'vendita', 'area', 'verde',
           'condominio', 'commenti', 'manutenzione', 'divisione', 'ponteggi',
           'protezione', 'dati', 'clienti', 'direzionali', 'tel', 'sei',
           'delegato', 'amministratore', 'titolare', 'direttore', 'presidente',
           'generale', 'commerciale', 'consulente', 'capo', 'logistico',
           'dott', 'sig', 'ing', 'geom', 'rag', 'produzione', 'italia'}
NAME_RE = re.compile(r"^[A-ZÉÈÊÀÂ][a-zàâäéèêëîïôöùûüç]+(-[A-ZÉÈÊÀÂ][a-zàâäéèêëîïôöùûüç]+)?$")

def clean_person(name):
    if not name:
        return ''
    parts = ' '.join(name.strip().split()).split(' ')
    # drop leading title words: 'Magasin Nicolas Faure' -> 'Nicolas Faure'
    while len(parts) > 2 and parts[0].lower() in LEADING:
        parts = parts[1:]
    # drop a repeated trailing word: 'Giles Potter Giles' -> 'Giles Potter'
    if len(parts) == 3 and parts[0] == parts[2]:
        parts = parts[:2]
    if len(parts) == 3 and parts[1] == parts[2]:
        parts = parts[:2]
    # strip trailing title/place words, keep at most the first two name words
    while len(parts) > 2 and parts[-1].lower() in TRAILING:
        parts = parts[:-1]
    if len(parts) != 2:
        return ''
    if any(p.lower() in GENERIC or p.lower() in TRAILING for p in parts):
        return ''
    if not all(NAME_RE.match(p) for p in parts):
        return ''
    if parts[0] == parts[1]:
        return ''
    if len(parts[1]) < 3:  # incomplete surname
        return ''
    return ' '.join(parts)

# ---------- grading ----------
TYPE_GRADE = {
    '经销商/代理商': 'A', '小挖经销商': 'A', '叉车经销商': 'A',
    '装载机经销商': 'A', '工程机械经销商': 'A',
    '机械租赁公司': 'B', '工程机械租赁': 'B',
    '园林绿化公司': 'C', '建筑/工程公司': 'D', '农机经销商': 'E',
}
GOOGLE_F = {'discount store', 'supermarket', 'grocery store', 'restaurant',
            'hotel', 'clothing store', 'bakery', 'pharmacy', 'furniture store',
            'department store', 'convenience store', 'drugstore', 'electronics store',
            'supermarché', 'hôtel', 'boulangerie', 'pharmacie', 'magasin de meubles'}
GOOGLE_GRADE = {
    'forklift dealer': 'A', 'construction machine dealer': 'A',
    'construction equipment supplier': 'A', 'industrial equipment supplier': 'A',
    'material handling equipment supplier': 'A', 'do-it-yourself shop': 'A',
    'lawn mower store': 'A', 'garden machinery supplier': 'A',
    'lawn mower repair service': 'A', 'tool store': 'A', 'hardware store': 'A',
    'plant and machinery hire': 'B', 'tool rental service': 'B',
    'equipment rental agency': 'B', 'crane rental agency': 'B',
    'landscaper': 'C', 'gardener': 'C', 'landscape designer': 'C',
    'landscape architect': 'C', 'garden center': 'C', 'lawn care service': 'C',
    'tree service': 'C',
    'construction company': 'D', 'building firm': 'D', 'general contractor': 'D',
    'home builder': 'D',
    'farm equipment supplier': 'E', 'tractor dealer': 'E',
}

def fallback_grade(x):
    gc = (x['谷歌类别'] or '').strip().lower()
    if gc in GOOGLE_F:
        return 'F', '谷歌分类为无关行业'
    g = TYPE_GRADE.get(x['客户类型'])
    if g:
        return g, '客户类型'
    if gc in GOOGLE_GRADE:
        return GOOGLE_GRADE[gc], '谷歌分类'
    return '', '未分级'

def tier_of(g):
    if g == 'A': return 'A级-优先联系'
    if g in ('B', 'E'): return 'B级-第二批'
    if g in ('C', 'D'): return 'C级-低匹配'
    if g == 'F': return 'D级-已剔除'
    return '未分级'

CAT_CN = {'mower': '割草机/园林机械', 'excavator': '小挖', 'forklift': '叉车',
          'loader': '装载机', 'machinery_dealer': '工程机械经销',
          'rental': '租赁', 'landscaping': '园林绿化', 'construction': '建筑工程',
          'agriculture': '农机', 'irrelevant': '与机械无关'}

def products(x):
    wc = (x.get('网站分类', '') or '')
    ct = x['客户类型']; gc = (x['谷歌类别'] or '').lower(); src = (x['来源搜索'] or '').lower()
    out = []
    if ('mower' in wc or 'mower' in src or 'garden-machinery' in src
            or gc in ('lawn mower store', 'lawn mower repair service',
                      'garden machinery supplier', 'garden center')):
        out.append('遥控割草机')
    if 'excavator' in wc or ct == '小挖经销商' or 'miniescavatore' in src or 'excavator' in src:
        out.append('小挖')
    if 'forklift' in wc or ct == '叉车经销商' or 'carrello' in src or 'forklift' in src:
        out.append('叉车')
    if 'loader' in wc or ct == '装载机经销商' or 'pala' in src or 'loader' in src:
        out.append('装载机')
    return out

def opener(x):
    prods = x['_products']
    lead = prods[0] if prods else '遥控割草机'
    prod_en = {'遥控割草机': 'remote-controlled tracked flail mowers',
               '小挖': 'mini excavators under 3 tonnes',
               '叉车': 'forklifts under 3 tonnes (fuel and pure electric)',
               '装载机': 'wheel loaders under 3 tonnes'}[lead]
    tier = x['_tier']
    if tier == 'A级-优先联系':
        ask = "Are you interested in dealership or distribution in Italy?"
    elif tier == 'B级-第二批':
        ask = "Would they fit your hire fleet in Italy?"
    elif tier == 'C级-低匹配':
        ask = ("Would they help with steep slopes, orchards and tall brush "
               "where an operator should not walk into the area?")
    else:
        ask = "Would they fit your groundworks and site work in Italy?"
    return (f"Hello, I'm David from HENBO Machinery in China. We manufacture "
            f"{prod_en}. {ask} "
            f"I can send a short field video and specifications on request.")

with open(SRC, encoding='utf-8-sig', newline='') as f:
    rows = list(csv.DictReader(f))

# name groups for 网点数
name_groups = defaultdict(list)
for x in rows:
    name_groups[x['客户名称'].strip().lower()].append(x)
phone_groups = defaultdict(list)
for x in rows:
    ph = re.sub(r'\D', '', x['电话'] or '')
    if ph:
        phone_groups[ph].append(x)

n_junk_any = n_junk_only = n_boels_junk = 0
for x in rows:
    raw_all = x['全部邮箱'] or ''
    parts = [p.strip() for p in raw_all.split(';') if p.strip()]
    had_junk = any(is_junk_email(p) for p in parts)
    if had_junk:
        n_junk_any += 1
        if 'boels' in (x['网站'] or '').lower() or 'boels' in x['客户名称'].lower():
            n_boels_junk += 1
    cleaned = clean_email_list(raw_all)
    if raw_all and not cleaned:
        n_junk_only += 1
    x['_全部邮箱'] = cleaned
    x['_主邮箱'] = cleaned.split(';')[0].strip() if cleaned else ''
    x['_had_junk'] = had_junk
    # person
    x['_联系人'] = clean_person(x['联系人'])
    # grade
    wg = (x['建议分级'] or '').strip()
    if wg:
        x['_最终分级'] = wg
        x['_判断来源'] = '读官网（网站分类：' + (x['网站分类'] or '无') + '）'
        x['_from_website'] = True
    else:
        g, why = fallback_grade(x)
        x['_最终分级'] = g
        x['_判断来源'] = f"未读到官网，按{why}补分级（网站：{x['网站可达']}）"
        x['_from_website'] = False
    x['_tier'] = tier_of(x['_最终分级'])
    x['_products'] = products(x)
    # city from first FR-<city>- segment in 来源搜索
    m = re.search(r'IT-([A-Za-z]+)-', x['来源搜索'] or '')
    x['_城市'] = m.group(1) if m else ''
    x['_网点数'] = len(name_groups[x['客户名称'].strip().lower()])
    # email tip
    if x['_had_junk'] and not cleaned:
        if n_boels_junk and 'boels' in (x['网站'] or '').lower():
            x['_邮箱提示'] = '原邮箱为假邮箱（图片文件名），已清空；Boels为连锁，总部邮箱待核实'
        else:
            x['_邮箱提示'] = '原邮箱为假邮箱（图片文件名/追踪地址），已清空'
    elif x['_had_junk']:
        x['_邮箱提示'] = '已去掉假邮箱（图片文件名/追踪地址），保留真邮箱'
    elif cleaned:
        n = len([p for p in cleaned.split(';') if p.strip()])
        x['_邮箱提示'] = f"共{n}个邮箱" if n > 1 else ''
    elif x['网站可达'] == '打不开':
        x['_邮箱提示'] = '网站打不开，未读到邮箱'
    elif x['网站可达'] == '无网站':
        x['_邮箱提示'] = '无网站，未读到邮箱'
    else:
        x['_邮箱提示'] = '官网未公开邮箱'
    # actual business
    wc = (x['网站分类'] or '').strip()
    if wc:
        x['_实际业务'] = '、'.join(CAT_CN.get(c.strip(), c.strip()) for c in wc.split(','))
    else:
        x['_实际业务'] = f"未读到官网内容，按客户类型判断：{x['客户类型']}（{x['谷歌类别']}）"
    # match / entry point
    prods = '、'.join(x['_products']) if x['_products'] else '产品线未明确'
    x['_匹配判断'] = f"{x['_tier'][:2]}；{x['客户类型']}，适合产品：{prods}"
    if x['_最终分级'] == 'F':
        x['_匹配判断'] += '；与机械无关，剔除'
    x['_开场白'] = opener(x)
    x['_备注'] = ''
    if x.get('删除原因'):
        x['_备注'] = x['删除原因']
    if x['_had_junk'] and cleaned:
        x['_备注'] = (x['_备注'] + '；' if x['_备注'] else '') + '注意：原邮箱里混有假地址，已去掉'

def stats(sub):
    return (len(sub),
            sum(1 for x in sub if x['_主邮箱'].strip()),
            sum(1 for x in sub if x['号码类型'] == '手机'))

tiers_order = ['A级-优先联系', 'B级-第二批', 'C级-低匹配', 'D级-已剔除']
for x in rows:
    pass
tier_rows = {t: [x for x in rows if x['_tier'] == t] for t in tiers_order}
total_stats = stats(rows)
ab = [x for x in rows if x['_最终分级'] in ('A', 'B')]
prod_names = ['遥控割草机', '小挖', '装载机', '叉车']
prod_rows = {p: [x for x in ab if p in x['_products']] for p in prod_names}
# duplicates: same name OR same phone
dup_keys = set()
for k, v in name_groups.items():
    if len(v) > 1:
        for x in v: dup_keys.add(id(x))
for k, v in phone_groups.items():
    if len(v) > 1:
        for x in v: dup_keys.add(id(x))
dup_rows = []
for x in rows:
    if id(x) in dup_keys:
        reasons = []
        if len(name_groups[x['客户名称'].strip().lower()]) > 1: reasons.append('同名')
        ph = re.sub(r'\D', '', x['电话'] or '')
        if ph and len(phone_groups[ph]) > 1: reasons.append('同电话')
        x['_重复原因'] = '、'.join(reasons)
        dup_rows.append(x)

print(f"total {len(rows)} email_cleaned {total_stats[1]} mobile {total_stats[2]}")
print(f"junk_any {n_junk_any} junk_only_empty {n_junk_only} boels_junk {n_boels_junk}")
print(f"person_cleaned {sum(1 for x in rows if x['_联系人'])}")
print(f"website_graded {sum(1 for x in rows if x['_from_website'])} fallback {sum(1 for x in rows if not x['_from_website'])}")
print("final grades", dict(collections.Counter(x['_最终分级'] for x in rows)))
for t in tiers_order: print(t, stats(tier_rows[t]))
for p in prod_names: print(p, stats(prod_rows[p]))
print("A+B", len(ab), "no product", sum(1 for x in ab if not x['_products']))
print("dup rows", len(dup_rows))
print("clean persons:", [x['_联系人'] for x in rows if x['_联系人']])

# cleaned CSV for audit trail
with open(CLEANED, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['客户名称', '最终分级', '分档', '客户类型', '主邮箱', '全部邮箱',
                '联系人（待核对）', '电话', '号码类型', '网站', '网站可达',
                '判断来源', '邮箱提示'])
    for x in rows:
        w.writerow([x['客户名称'], x['_最终分级'], x['_tier'], x['客户类型'],
                    x['_主邮箱'], x['_全部邮箱'], x['_联系人'], x['电话'],
                    x['号码类型'], x['网站'], x['网站可达'], x['_判断来源'],
                    x['_邮箱提示']])

# ---------- workbook ----------
wb = Workbook()
title_font = Font(name='Calibri', size=14, bold=True)
header_fill = PatternFill('solid', fgColor='D9E2F3')
header_font = Font(name='Calibri', size=11, bold=True)
thin = Side(style='thin', color='B0B0B0')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical='center')
center = Alignment(horizontal='center', vertical='center', wrap_text=True)

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = center; cell.border = border

def detail_columns():
    return ['公司名称', '最终分级', '客户类型', '谷歌分类', '适合产品', '实际业务（读官网后）',
            '匹配判断/切入点', '主邮箱', '全部邮箱', '邮箱提示', '联系人（负责人，待核对）',
            '电话', '号码类型', 'WhatsApp', 'WhatsApp校验', '独立站', '地址', '城市（搜索城市）',
            'Facebook', 'Instagram', 'LinkedIn', 'YouTube', 'TikTok', 'X/Twitter',
            '评分', '评论数', '来源搜索词', '网点数', '判断来源', '开发信英文开场白', '备注']

def write_detail_sheet(ws, data):
    cols = detail_columns()
    ws.append(cols)
    style_header(ws, 1, len(cols))
    for x in data:
        ws.append([
            x['客户名称'], x['_最终分级'], x['客户类型'], x['谷歌类别'],
            '、'.join(x['_products']), x['_实际业务'], x['_匹配判断'],
            x['_主邮箱'], x['_全部邮箱'], x['_邮箱提示'], x['_联系人'],
            x['电话'], x['号码类型'], x['WhatsApp'], x['WhatsApp校验'],
            x['网站'], x['地址'], x['_城市'],
            x['Facebook'], x['Instagram'], x['LinkedIn'], x['YouTube'],
            x['TikTok'], x['X/Twitter'], x['评分'], x['评论数'],
            x['来源搜索'], x['_网点数'], x['_判断来源'], x['_开场白'],
            (('重复原因：' + x['_重复原因'] + '；') if x.get('_重复原因') else '') + x['_备注'],
        ])
    widths = [30, 10, 16, 26, 18, 30, 34, 30, 34, 30, 18, 18, 10, 32, 22, 36, 40, 14,
              30, 30, 30, 30, 30, 30, 8, 8, 30, 8, 34, 60, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=2, column=1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{1 + len(data)}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

# --- 统计 sheet ---
ws = wb.active; ws.title = '统计'
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.column_dimensions['A'].width = 22; ws.column_dimensions['B'].width = 14
ws.column_dimensions['C'].width = 16; ws.column_dimensions['D'].width = 24
ws.merge_cells('A1:D1'); c = ws['A1']
c.value = 'HENBO客户审核·意大利·2092家（谷歌地图+读官网）'; c.font = title_font
c.alignment = Alignment(vertical='center')
ws.merge_cells('A2:D2')
ws['A2'] = (f"底表：IT_enriched.csv（2092行，读官网完成2026-11-08约01:39 UTC，"
            f"清洗后邮箱{total_stats[1]}家，负责人（待核对）"
            f"{sum(1 for x in rows if x['_联系人'])}家，手机号{total_stats[2]}家）。"
            f"德国/英国/法国表格式，共12页。")
ws['A2'].alignment = wrap; ws.row_dimensions[2].height = 30
r = 4
ws.cell(row=r, column=1, value='项目'); ws.cell(row=r, column=2, value='家数')
ws.cell(row=r, column=3, value='其中有邮箱'); ws.cell(row=r, column=4, value='其中有手机号（可试WhatsApp）')
style_header(ws, r, 4); r += 1
for t in tiers_order:
    n, e, m = stats(tier_rows[t])
    ws.cell(row=r, column=1, value=t); ws.cell(row=r, column=2, value=n)
    ws.cell(row=r, column=3, value=e); ws.cell(row=r, column=4, value=m)
    for col in range(1, 5):
        ws.cell(row=r, column=col).border = border; ws.cell(row=r, column=col).alignment = center
    r += 1
ws.cell(row=r, column=1, value='合计（去重后）'); ws.cell(row=r, column=2, value=total_stats[0])
ws.cell(row=r, column=3, value=total_stats[1]); ws.cell(row=r, column=4, value=total_stats[2])
for col in range(1, 5):
    ws.cell(row=r, column=col).font = Font(bold=True)
    ws.cell(row=r, column=col).border = border; ws.cell(row=r, column=col).alignment = center
r += 2
ws.merge_cells(f'A{r}:D{r}'); ws.cell(row=r, column=1, value='产品线（A+B）').font = Font(bold=True, size=12); r += 1
ws.cell(row=r, column=1, value='产品线'); ws.cell(row=r, column=2, value='家数')
ws.cell(row=r, column=3, value='其中有邮箱'); ws.cell(row=r, column=4, value='其中有手机号')
style_header(ws, r, 4); r += 1
for p in prod_names:
    n, e, m = stats(prod_rows[p])
    ws.cell(row=r, column=1, value=p); ws.cell(row=r, column=2, value=n)
    ws.cell(row=r, column=3, value=e); ws.cell(row=r, column=4, value=m)
    for col in range(1, 5):
        ws.cell(row=r, column=col).border = border; ws.cell(row=r, column=col).alignment = center
    r += 1
r += 1
# ---- computed facts for the notes ----
grade_counter = collections.Counter(x['_最终分级'] for x in rows)
n_website = sum(1 for x in rows if x['_from_website'])
n_fallback = len(rows) - n_website
reach_counter = collections.Counter(x['网站可达'] for x in rows)
n_raw_person = sum(1 for x in rows if (x['联系人'] or '').strip())
n_clean_person = sum(1 for x in rows if x['_联系人'])
n_landline = sum(1 for x in rows if x['号码类型'] == '座机')
n_nophone = sum(1 for x in rows if not x['号码类型'])
soc_counts = {k: sum(1 for x in rows if (x[k] or '').strip())
              for k in ['Facebook', 'Instagram', 'LinkedIn', 'YouTube', 'TikTok', 'X/Twitter']}
dup_top = sorted(((k, len(v)) for k, v in name_groups.items() if len(v) > 1),
                 key=lambda kv: -kv[1])[:2]


def _name_for(k):
    for x in rows:
        if x['客户名称'].strip().lower() == k:
            return x['客户名称']
    return k


dup_top_txt = '、'.join(f"{_name_for(k)} {n}家" for k, n in dup_top) if dup_top else '无'
n_cities = len({x['_城市'] for x in rows if x['_城市']})
boels_txt = (f"Boels意大利{n_boels_junk}家分店的“邮箱”全是图片文件名，清洗后变空，"
             f"需另核实总部邮箱，不要用那个假地址发信。" if n_boels_junk else '')
ws.merge_cells(f'A{r}:D{r}')
ws.cell(row=r, column=1, value=f"原分级对照（读官网{n_website}家+按类型补分级{n_fallback}家）："
      f"A {grade_counter['A']}家，"
      f"B {grade_counter['B']}家，"
      f"C {grade_counter['C']}家，"
      f"D {grade_counter['D']}家，"
      f"E {grade_counter['E']}家，"
      f"F {grade_counter['F']}家。"
      f"四档映射与德国/英国表相同：A级=A；B级=B+E；C级=C+D；D级=F。").font = Font(size=10)
ws.cell(row=r, column=1).alignment = wrap; ws.row_dimensions[r].height = 36; r += 2
ws.merge_cells(f'A{r}:D{r}'); ws.cell(row=r, column=1, value='说明：').font = Font(bold=True, size=12); r += 1
notes = [
    f'1. 意大利没有“州”，本表用“城市（搜索城市）”代替，即抓取时用的城市（共{n_cities}个城市，新产品搜索用的12个大城市：罗马、米兰、那不勒斯、都灵、巴勒莫、热那亚、博洛尼亚、佛罗伦萨、巴里、卡塔尼亚、威尼斯、维罗纳）；地址列里有完整地址和邮编，可再核对省。',
    (f'2. 假邮箱已清洗：{n_junk_any}行邮箱里混有假地址（图片文件名如 logo@2x.png、追踪地址如 @sentry.io），其中{n_junk_only}行只有假邮箱，清洗后变空邮箱；清洗后全表有邮箱{total_stats[1]}家。{boels_txt}'
     if n_junk_any else
     f'2. 假邮箱：按德国/英国/法国表同样规则检查过，意大利表里没有发现假地址（图片文件名、追踪地址），全表有邮箱{total_stats[1]}家。'),
    f'3. 负责人名字很少且未核实：网站上只抓到{n_raw_person}个名字，清洗掉头衔词/碎片后剩{n_clean_person}个（见各页“联系人”列），其余公司官网没有公开负责人名字。意大利没有德国/法国那种法定公示页上强制公开负责人名字的做法，这是正常情况，不是漏抓。发信请用公司邮箱+负责人名字（如有），不要编私人邮箱。',
    f'4. 分级缺口已补：读官网只分出{n_website}家（网站能打开{reach_counter["是"]}家，{reach_counter["打不开"]}家打不开，{reach_counter["无网站"]}家无网站），其余{n_fallback}家按客户类型/谷歌分类补分级，判断来源列里写明是“读官网”还是“按类型补分级”，补分级的分级依据弱于读官网的，联系前可再抽查。',
    f'5. 手机号{total_stats[2]}家，只按意大利3开头判断为手机，WhatsApp列给的是 wa.me 链接，校验列写明“未验证（需在WhatsApp里确认是否注册）”——自动校验做不到，见德国/英国/法国表同样说明。座机{n_landline}家，WhatsApp不适用；{n_nophone}家无电话。',
    f'6. 社交账号只统计官网链接到的：Facebook {soc_counts["Facebook"]}家、Instagram {soc_counts["Instagram"]}家、LinkedIn {soc_counts["LinkedIn"]}家、YouTube {soc_counts["YouTube"]}家、TikTok {soc_counts["TikTok"]}家、X/Twitter {soc_counts["X/Twitter"]}家，官网没链接的为空，不代表没有账号。',
    '7. 开发信英文开场白是按分档/产品生成的模板，没有编价格、起订量、交货期，发之前请按实际情况改。HENBO小挖目前还没上市，先建立联系、收集需求，等样机/参数确定后再报价（与德国/英国/法国表同样提醒）。',
    f'8. 重复条目多为连锁分店同名不同地址/电话（如 {dup_top_txt}），不要只按同名删除；网点数列标了同名门店数。',
    f'9. 产品线按现有资料判断：一家可同时卖多种产品。A+B共{len(ab)}家，其中{sum(1 for x in ab if not x["_products"])}家在现有资料里看不出明确产品线，未强行归类，未计入产品线四行（与德国/英国/法国表同样做法，德国是526家，英国是776家，法国是647家）。',
]
for n in notes:
    ws.merge_cells(f'A{r}:D{r}'); cell = ws.cell(row=r, column=1, value=n)
    cell.alignment = wrap; cell.font = Font(size=10); ws.row_dimensions[r].height = 48; r += 1

def add_sheet(name, data):
    s = wb.create_sheet(name)
    write_detail_sheet(s, data)
    return s

add_sheet('A级-优先联系', tier_rows['A级-优先联系'])
add_sheet('B级-第二批', tier_rows['B级-第二批'])
add_sheet('遥控割草机(A+B)', prod_rows['遥控割草机'])
add_sheet('小挖(A+B)', prod_rows['小挖'])
add_sheet('装载机(A+B)', prod_rows['装载机'])
add_sheet('叉车(A+B)', prod_rows['叉车'])
add_sheet('C级-低匹配', tier_rows['C级-低匹配'])
add_sheet('D级-已剔除', tier_rows['D级-已剔除'])
add_sheet('重复条目', dup_rows)
add_sheet('全部(去重后)', rows)

ks = wb.create_sheet('意大利语关键词清单')
ks.column_dimensions['A'].width = 16; ks.column_dimensions['B'].width = 28
ks.column_dimensions['C'].width = 28; ks.column_dimensions['D'].width = 34
ks.merge_cells('A1:D1'); ks['A1'] = '意大利语关键词清单（本次实际使用的搜索词）'; ks['A1'].font = title_font
ks.append([]); ks.append(['产品', '意大利语关键词', '搜索词（地图抓取用）', '说明'])
style_header(ks, 3, 4)
kw_rows = [
    ['遥控割草机', 'tagliaerba / macchine da giardino', 'mower-dealer / garden-machinery', '割草机经销商、园林机械店'],
    ['小挖', 'mini escavatore', 'miniescavatore', '3吨以下小挖'],
    ['叉车', 'carrello elevatore', 'carrello', '叉车，含燃油和纯电，3吨以下'],
    ['装载机', 'pala gommata', 'pala', '装载机，3吨以下'],
    ['通用机械', 'macchine movimento terra', 'movimento', '工程机械经销商'],
    ['租赁', 'noleggio macchine edili', 'noleggio / rental', '工程机械租赁、机械租赁'],
    ['园林', 'giardiniere / manutenzione del verde', 'landscaping', '园林绿化公司'],
    ['建筑', 'edilizia / costruzioni', 'construction', '建筑/工程公司'],
]
for row in kw_rows: ks.append(row)
for row in ks.iter_rows(min_row=4, max_row=3 + len(kw_rows), max_col=4):
    for cell in row: cell.border = border; cell.alignment = wrap
ks.append([]); ks.append(['注：地图抓取用意大利语搜索词+意大利城市名组合；新产品（小挖/叉车/装载机/工程机械/租赁）在罗马、米兰、那不勒斯、都灵、巴勒莫、热那亚、博洛尼亚、佛罗伦萨、巴里、卡塔尼亚、威尼斯、维罗纳12个城市各跑5类，共60条搜索。'])
ks.merge_cells(f'A{4 + len(kw_rows)}:D{4 + len(kw_rows)}')

wb.save(OUT)
print("saved", OUT)
