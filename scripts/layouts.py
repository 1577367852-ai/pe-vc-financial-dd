"""Bounded layout adapters; ambiguous/continued metadata remains reviewable evidence."""
import re
from pathlib import Path

META = ('主体','期间','范围','币种','版本','截止日','单位')
PERIOD = re.compile(r'^(?:20\d{2}(?:年度|年|-[HQ][1-4]|-\d\d-\d\d)?|FY20\d{2})$')


def matrix_records(wb, cache, path, digest):
    from ingest import record, ALIASES, limited_formula
    rows, handled = [], set()
    for ws in wb:
        meta, header = {}, None
        for cells in ws:
            for a,b in zip(cells,cells[1:]):
                if a.value in META and b.value is not None: meta[a.value]=str(b.value)
        for cells in ws:
            labels=[c for c in cells if c.value in ('科目','项目')]
            periods={c.column:str(c.value) for c in cells if PERIOD.fullmatch(str(c.value or ''))}
            if labels and periods:
                header=(labels[0].column,periods,next((c.column for c in cells if c.value=='单位'),None))
                handled.add(ws.title)
                continue
            if header is None: continue
            label_col, periods, unit_col=header
            label=ws.cell(cells[0].row,label_col).value
            if label not in ALIASES: continue
            unit=ws.cell(cells[0].row,unit_col).value if unit_col else meta.get('单位')
            for col,period in periods.items():
                c=ws.cell(cells[0].row,col); value,status=c.value,'READ'
                formula=c.value if c.data_type=='f' else None
                if formula:
                    try: value,status=limited_formula(wb,ws.title,c.coordinate),'RECOMPUTED'
                    except (ValueError,ArithmeticError,KeyError): value,status=None,'UNSUPPORTED_FORMULA'
                m=dict(meta,期间=period)
                if meta.get('期间')!=period: m['截止日']=None
                src=dict(file=str(Path(path).resolve()),sha256=digest,sheet=ws.title,cell=c.coordinate,
                         label_cell=ws.cell(c.row,label_col).coordinate,period_column=col,
                         sheet_state=ws.sheet_state,merged_ranges=[str(r) for r in ws.merged_cells.ranges])
                r=record(label,value,unit,m,src,status,formula,cache[ws.title][c.coordinate].value if formula else None)
                r['original_value']=c.value; rows.append(r)
    return rows,handled


def continued_text(lines, path, digest, page, inherited=None, boxes=None, ocr=False):
    from ingest import record, parse_text_rows, ALIASES
    meta=dict(inherited or {})
    explicit={}
    for line in lines:
        m=re.match(r'^(主体|期间|范围|币种|版本|截止日|单位)\s*[:：]\s*(.+)$',line.strip())
        if m: explicit[m[1]]=m[2].strip()
    # A changed entity resets all inherited identity fields.
    if '主体' in explicit and explicit['主体']!=meta.get('主体'):meta={}
    meta.update(explicit)
    # Prefix only metadata missing on this page; preserve actual line numbers afterwards.
    prefix=[f'{k}: {v}' for k,v in meta.items() if k in META and k not in explicit]
    records=parse_text_rows(prefix+lines,path,digest,page,ocr=ocr)
    for r in records:
        line=r['source']['line']-len(prefix);r['source']['line']=line
        r['source']['bbox']=(boxes or {}).get(line)
    header=meta.get('_table_header')
    for i,line in enumerate(lines,1):
        parts=[x.strip() for x in line.split('|')]
        if parts[0] in ('科目','项目') and any(PERIOD.fullmatch(x) for x in parts[1:]):
            header=parts;meta['_table_header']=parts;continue
        if not header or parts[0] not in ALIASES or len(parts)!=len(header):continue
        unit=parts[header.index('单位')] if '单位' in header else meta.get('单位')
        for j,p in enumerate(header):
            if not PERIOD.fullmatch(p):continue
            m=dict(meta,期间=p)
            if meta.get('期间')!=p:m['截止日']=None
            src=dict(file=str(Path(path).resolve()),sha256=digest,physical_page=page,line=i,
                     original_line=line,column=j+1,bbox=(boxes or {}).get(i),bbox_status='VERIFY_ON_PAGE')
            records.append(record(parts[0],parts[j],unit,m,src,'NEEDS_REVIEW' if ocr else 'READ'))
    if prefix:
        for r in records:
            r['source']['inherited_metadata']=dict(meta)
            r['candidate_value']=r['value'];r['value']=None;r['status']='NEEDS_REVIEW'
    return records,meta
