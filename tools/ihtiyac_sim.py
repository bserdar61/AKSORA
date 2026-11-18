import openpyxl,re,sys
wb=openpyxl.load_workbook('up62/AKSORA.xlsm',read_only=True)
ILLER=['ARTVIN','TRABZON','GUMUSHANE','GIRESUN','RIZE']
TR=str.maketrans('İıŞĞÜÖÇşğüöç','IISGUOCSGUOC')
def norm(t): return str(t or '').upper().translate(TR).strip()
def soz(t):
    t=norm(t); out=[];cur='';on=0
    for ch in t+' ':
        tur=1 if 'A'<=ch<='Z' else 2 if ch.isdigit() else 0
        if tur!=on and cur:
            if on==2: cur=cur.lstrip('0') or '0'
            out.append(cur);cur=''
        if tur: cur+=ch
        on=tur
    return ' '+' '.join(out)+' '
def uygun(ad,parti):
    t=' '+norm(ad)+' '
    return not any(k in t for k in ['HURDA','DEMONTAJ','ARIZA',' DM ']) and 'HURDA' not in norm(parti)
d={}
for il in ILLER:
    src=['CDP_STOK','EPP_STOK'] if il=='ARTVIN' else ['_AKLIVE_%s_CDP'%il,'_AKLIVE_%s_EPP'%il]
    for s in src:
        for r in wb[s].iter_rows(min_row=7,max_col=11,values_only=True):
            dp=str(r[1] or '').upper()
            if not re.fullmatch(r'D\d{3}',dp): continue
            mal=str(r[6]).split('.')[0] if r[6] else ''
            x=float(r[8] or 0)
            if not mal or x<=0: continue
            o=d.setdefault(mal,[r[7],r[9],soz(r[7]),{},[],norm(r[7])])
            u=uygun(r[2],r[5])
            o[4].append((il,dp,r[2],str(r[5]),x,float(r[10] or 0),u))
            if u: o[3][il]=o[3].get(il,0)+x
def esles(m,sz):
    ok=False
    for x in sz:
        if x[0].isdigit():
            if ' '+x+' ' not in m: return False
        elif ' '+x not in m:
            if x=='TRAFO' and ' TRF ' in m: pass
            elif x=='TRF' and ' TRAFO' in m: pass
            else: return False
        ok=True
    return ok
def ara(q):
    qq=q.replace(' ','')
    if len(qq)>=10 and qq[:10].isdigit() and qq[:10] in d: return [qq[:10]]
    rak=len(qq)>=3 and qq.isdigit(); sz=soz(q).split(); qn=norm(q)
    return [k for k,o in d.items() if (qq in k if rak else (esles(o[2],sz) or (len(qn)>=3 and qn in o[5])))]
il='ARTVIN'
for q,n in [('nh00 63',50),('6161',100),('50 kva trafo',3),('bant izole',500)]:
    es=ara(q); print('\n###',q,n,'->',len(es),'malzeme')
    rows=[]
    for m in es:
        o=d[m]; own=o[3].get(il,0); oth=sum(v for k,v in o[3].items() if k!=il)
        p=own+oth+(2e15 if own>=n else 1e15 if own+oth>=n else 0)
        rows.append((p,m,o[0],o[1],own,{k:o[3].get(k,0) for k in ILLER[1:]}))
    rows.sort(reverse=True)
    for r in rows[:6]: print('  ',r[1],r[2],r[3],'kendi',r[4],r[5])
