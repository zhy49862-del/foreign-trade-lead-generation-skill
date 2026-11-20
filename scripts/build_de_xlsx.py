import csv, re, collections
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = Path.home()/"workspace/goals/rc-mower-export-lead-prospecting/files/DE_supplemented_batch1.csv"
OUT = Path.home()/"workspace/goals/rc-mower-export-lead-prospecting/files/HENBO德国客户审核_2705家.xlsx"

def is_mobile(ph):
    s=re.sub(r'\D','', ph or '')
    if s.startswith('49'): s=s[2:]
    elif s.startswith('0'): s=s[1:]
    return s[:2] in ('15','16','17')

def products(x):
    wc=(x.get('网站分类','') or '')
    ct=x['客户类型']; gc=x['谷歌类别']; src=x['来源搜索']
    out=set()
    if ct=='小挖经销商' or 'excavator' in wc: out.add('小挖')
    if ct=='叉车经销商' or 'forklift' in wc: out.add('叉车')
    if ct=='装载机经销商' or 'loader' in wc: out.add('装载机')
    if 'mower' in wc or gc in ('Lawn mower store','Lawn mower repair service','Garden machinery supplier','Garden center') or 'mower' in src or 'garden-machinery' in src:
        out.add('遥控割草机')
    return out

with open(SRC, encoding='utf-8-sig', newline='') as f:
    rows=list(csv.DictReader(f))

# Mapping to image's A/B/C/D
def tier(x):
    g=x['最终分级']
    if g=='A': return 'A级-优先联系'
    if g in ('B','E'): return 'B级-第二批'
    if g in ('C','D'): return 'C级-低匹配'
    if g=='F': return 'D级-已剔除'
    return '未分级'

for x in rows:
    x['_tier']=tier(x)
    x['_mobile']='是' if (x['电话'].strip() and is_mobile(x['电话'])) else ('否' if x['电话'].strip() else '')
    x['_products']=products(x)
    x['_备注']=''
    notes=[]
    if x.get('补充说明','').strip(): notes.append(x['补充说明'].strip())
    if x.get('补充来源','').strip(): notes.append('来源：'+x['补充来源'].strip())
    x['_备注']='；'.join(notes)

def stats(sub):
    return (len(sub), sum(1 for x in sub if x['邮箱'].strip()), sum(1 for x in sub if x['_mobile']=='是'))

tiers_order=['A级-优先联系','B级-第二批','C级-低匹配','D级-已剔除']
tier_rows={t:[x for x in rows if x['_tier']==t] for t in tiers_order}
total_stats=stats(rows)
ab=[x for x in rows if x['最终分级'] in ('A','B')]
prod_names=['遥控割草机','小挖','装载机','叉车']
prod_rows={p:[x for x in ab if p in x['_products']] for p in prod_names}

# Duplicates: same name (lower, stripped) OR same phone digits
from collections import defaultdict
name_groups=defaultdict(list); phone_groups=defaultdict(list)
for x in rows:
    name_groups[x['客户名称'].strip().lower()].append(x)
    ph=re.sub(r'\D','', x['电话'] or '')
    if ph: phone_groups[ph].append(x)
dup_keys=set()
dup_rows=[]
seen=set()
for k,v in name_groups.items():
    if len(v)>1:
        for x in v: dup_keys.add(id(x))
for k,v in phone_groups.items():
    if len(v)>1:
        for x in v: dup_keys.add(id(x))
for x in rows:
    if id(x) in dup_keys:
        reasons=[]
        if len(name_groups[x['客户名称'].strip().lower()])>1: reasons.append('同名')
        ph=re.sub(r'\D','', x['电话'] or '')
        if ph and len(phone_groups[ph])>1: reasons.append('同电话')
        x['_重复原因']='、'.join(reasons)
        dup_rows.append(x)
print(f"total {len(rows)} email {total_stats[1]} mobile {total_stats[2]}")
for t in tiers_order: print(t, stats(tier_rows[t]))
for p in prod_names: print(p, stats(prod_rows[p]))
print("A+B", len(ab), "no product", sum(1 for x in ab if not x['_products']))
print("dup rows", len(dup_rows))

wb=Workbook()
# styles
title_font=Font(name='Calibri', size=14, bold=True)
header_fill=PatternFill('solid', fgColor='D9E2F3')
header_font=Font(name='Calibri', size=11, bold=True)
thin=Side(style='thin', color='B0B0B0')
border=Border(left=thin,right=thin,top=thin,bottom=thin)
wrap=Alignment(wrap_text=True, vertical='center')
center=Alignment(horizontal='center', vertical='center', wrap_text=True)

def style_header(ws, row, ncols):
    for c in range(1, ncols+1):
        cell=ws.cell(row=row, column=c)
        cell.fill=header_fill; cell.font=header_font; cell.alignment=center; cell.border=border

def detail_columns():
    return ['客户名称','最终分级','客户类型','邮箱','联系人（经理/老板）','电话','是否手机号','网站','网站可达','地址','谷歌类别','评分','评论数','谷歌地图链接','来源搜索','备注']

def write_detail_sheet(ws, data, title_note=None):
    cols=detail_columns()
    ws.append([title_note] if title_note else cols)
    if title_note:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
        ws.cell(row=1,column=1).font=Font(bold=True, size=11)
        ws.append(cols)
        header_row=2
    else:
        header_row=1
    style_header(ws, header_row, len(cols))
    for x in data:
        ws.append([
            x['客户名称'], x['最终分级'], x['客户类型'], x['邮箱'], x['联系人'],
            x['电话'], x['_mobile'], x['网站'], x['网站可达'], x['地址'],
            x['谷歌类别'], x['评分'], x['评论数'], x['谷歌地图链接'], x['来源搜索'], x['_备注']
        ])
    widths=[28,10,16,28,18,18,10,38,10,38,24,8,8,30,28,50]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes=ws.cell(row=header_row+1, column=1)
    ws.auto_filter.ref=f"A{header_row}:{get_column_letter(len(cols))}{header_row+len(data)}"
    ws.sheet_properties.pageSetUpPr.fitToPage=True

# --- 统计 sheet ---
ws=wb.active; ws.title='统计'
ws.sheet_properties.pageSetUpPr.fitToPage=True
ws.column_dimensions['A'].width=22; ws.column_dimensions['B'].width=14; ws.column_dimensions['C'].width=16; ws.column_dimensions['D'].width=24
ws.merge_cells('A1:D1'); c=ws['A1']; c.value='HENBO客户审核·德国·2705家（谷歌地图+谷歌搜索补充）'; c.font=title_font; c.alignment=Alignment(vertical='center')
ws.merge_cells('A2:D2'); ws['A2']='底表：DE_supplemented_batch1.csv（2705行，邮箱963家，负责人511家），统计日期：2026-11-07。图片里的21家小挖测试只作格式参考，数字以本表为准。'; ws['A2'].alignment=wrap; ws.row_dimensions[2].height=30
r=4
ws.cell(row=r,column=1,value='项目'); ws.cell(row=r,column=2,value='家数'); ws.cell(row=r,column=3,value='其中有邮箱'); ws.cell(row=r,column=4,value='其中有手机号（可试WhatsApp）')
style_header(ws,r,4)
r+=1
for t in tiers_order:
    n,e,m=stats(tier_rows[t])
    ws.cell(row=r,column=1,value=t); ws.cell(row=r,column=2,value=n); ws.cell(row=r,column=3,value=e); ws.cell(row=r,column=4,value=m)
    for col in range(1,5): ws.cell(row=r,column=col).border=border; ws.cell(row=r,column=col).alignment=center
    r+=1
ws.cell(row=r,column=1,value='合计（去重后）'); ws.cell(row=r,column=2,value=total_stats[0]); ws.cell(row=r,column=3,value=total_stats[1]); ws.cell(row=r,column=4,value=total_stats[2])
for col in range(1,5): ws.cell(row=r,column=col).font=Font(bold=True); ws.cell(row=r,column=col).border=border; ws.cell(row=r,column=col).alignment=center
r+=2
ws.merge_cells(f'A{r}:D{r}'); ws.cell(row=r,column=1,value='产品线（A+B）').font=Font(bold=True, size=12); r+=1
ws.cell(row=r,column=1,value='产品线'); ws.cell(row=r,column=2,value='家数'); ws.cell(row=r,column=3,value='其中有邮箱'); ws.cell(row=r,column=4,value='其中有手机号')
style_header(ws,r,4); r+=1
for p in prod_names:
    n,e,m=stats(prod_rows[p])
    ws.cell(row=r,column=1,value=p); ws.cell(row=r,column=2,value=n); ws.cell(row=r,column=3,value=e); ws.cell(row=r,column=4,value=m)
    for col in range(1,5): ws.cell(row=r,column=col).border=border; ws.cell(row=r,column=col).alignment=center
    r+=1
r+=1
ws.merge_cells(f'A{r}:D{r}'); ws.cell(row=r,column=1,value='原最终分级（A-F）对照：A 1020家，B 470家，C 982家，D 80家，E 127家，F 26家。图片格式只分四档，本表映射为：A级=A；B级=B+E；C级=C+D；D级=F。原分级在“全部”页的“最终分级”列里完整保留。').font=Font(size=10); ws.cell(row=r,column=1).alignment=wrap; ws.row_dimensions[r].height=36; r+=2
ws.merge_cells(f'A{r}:D{r}'); ws.cell(row=r,column=1,value='说明：').font=Font(bold=True, size=12); r+=1
notes=[
'1. 本次底表是谷歌地图抓取的德国2705家，不是图片里21家小挖测试；图片只作为表格格式参考（统计表+按产品线分表+备注写经理名字）。',
'2. 邮箱和负责人已用浏览器指令方法补充：先读官网首页/联系页/公示页（Impressum，德国法律要求公开的公司信息页），再用谷歌搜索 site:域名 inurl:impressum、site:域名 inurl:kontakt 和“公司名+Geschäftsführer/ Inhaber”补缺。样品3家+批量161家已补，全表邮箱836→963家，负责人399→511家。',
'3. 德国手机号以 +49 15 / 16 / 17 开头，本表只按号码开头判断是否为手机号，需要在WhatsApp里逐个确认是否注册，不能只看号码就认定能加WhatsApp。',
'4. 重点提醒：补到的邮箱里有92行是连锁总部统一邮箱（Boels 76行、Würth 16行），负责人也是总部负责人，不是分店负责人，发信不要当成分店联系人。独立公司69行里只补到35家有邮箱。',
'5. 产品线按现有资料判断：一家可同时卖多种产品，所以四行合计大于A+B家数。小挖=小挖经销商或网站分类含excavator；叉车、装载机同理；遥控割草机=网站分类含mower或割草机/园林机械类店铺。另有526家A+B在现有资料里没有明确产品线，未强行归类，未计入四行。',
'6. 注意：HENBO小挖目前还没上市，先建立联系、收集需求，等样机/参数确定后再报价。',
'7. D级26家已标删除原因，未删除，等你确认再删；重复条目多为连锁分店同名不同地址/电话，不要只按同名删除。官网打不开时，已有联系方式需另行核实。',
]
for n in notes:
    ws.merge_cells(f'A{r}:D{r}'); cell=ws.cell(row=r,column=1,value=n); cell.alignment=wrap; cell.font=Font(size=10); ws.row_dimensions[r].height=36; r+=1

# --- detail sheets ---
def add_sheet(name, data):
    s=wb.create_sheet(name)
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
# 重复条目 with extra column? reuse detail + reason in 备注 prefix
for x in dup_rows:
    x['_备注']=('重复原因：'+x['_重复原因']+'；'+x['_备注']) if x['_备注'] else ('重复原因：'+x['_重复原因'])
add_sheet('重复条目', dup_rows)
add_sheet('全部(去重后)', rows)

# --- 德语关键词清单 ---
ks=wb.create_sheet('德语关键词清单')
ks.column_dimensions['A'].width=16; ks.column_dimensions['B'].width=28; ks.column_dimensions['C'].width=28; ks.column_dimensions['D'].width=34
ks.merge_cells('A1:D1'); ks['A1']='德语关键词清单（本次实际使用的搜索词）'; ks['A1'].font=title_font
headers=['产品','德语关键词','英语搜索词（地图抓取用）','说明']
ks.append([]); ks.append(headers); style_header(ks,3,4)
kw_rows=[
['遥控割草机','Rasenmäher / Gartengeräte','mower-dealer / garden-machinery','割草机经销商、园林机械店'],
['遥控割草机','Rasenmäher Händler','mower-dealer','割草机经销商'],
['小挖','Minibagger','minibagger','3吨以下小挖，德语核心词'],
['小挖','Minibagger Händler','minibagger','小挖经销商'],
['叉车','Gabelstapler','gabelstapler','叉车，含燃油和纯电，3吨以下'],
['叉车','Gabelstapler Händler','gabelstapler','叉车经销商'],
['装载机','Radlader','radlader','装载机，3吨以下'],
['装载机','Radlader Händler','radlader','装载机经销商'],
['通用机械','Baumaschinenhändler','baumaschinen','工程机械经销商'],
['通用机械','Baumaschinenverleih','baumaschinenverleih','工程机械租赁'],
['农机','Landmaschinen','agri-machinery','农机经销商'],
['园林','Garten / Landschaftsbau','landscaping','园林绿化公司'],
['租赁','Vermietung / Verleih','rental','机械租赁公司'],
]
for row in kw_rows: ks.append(row)
for row in ks.iter_rows(min_row=4, max_row=3+len(kw_rows), max_col=4):
    for cell in row: cell.border=border; cell.alignment=wrap
ks.append([]); ks.append(['注：地图抓取用的是英语搜索词+德国城市名组合；新产品（小挖/叉车/装载机）用的是上面德语词，在柏林、慕尼黑、汉堡、科隆、法兰克福、斯图加特、杜塞尔多夫、莱比锡、多特蒙德、埃森、汉诺威、纽伦堡12个城市各跑5类，共60条搜索。'])
ks.merge_cells(f'A{4+len(kw_rows)}:D{4+len(kw_rows)}')

wb.save(OUT)
print("saved", OUT)
