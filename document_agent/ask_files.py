"""Local, evidence-grounded Q&A helpers for research PDFs and Excel workbooks."""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any

import fitz
import pandas as pd
from PIL import Image
from .extraction import _ocr


@dataclass
class PaperPage:
    page: int
    text: str


def read_paper_pdf(name: str, content: bytes, languages: str='eng+hin') -> list[PaperPage]:
    if not name.lower().endswith('.pdf') or not content.startswith(b'%PDF-'):
        raise ValueError('Upload a valid PDF research paper.')
    doc=fitz.open(stream=content,filetype='pdf')
    if doc.page_count>200: raise ValueError('Research PDFs are limited to 200 pages.')
    pages=[]
    for i,p in enumerate(doc):
        text=p.get_text('text').strip()
        if len(text)<30:
            pix=p.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
            text=_ocr(Image.open(io.BytesIO(pix.tobytes('png'))),languages).strip()
        pages.append(PaperPage(i+1,text))
    doc.close()
    if not any(p.text for p in pages): raise ValueError('No embedded text was found. OCR this scanned PDF first, then ask questions about the extracted text.')
    return pages


def _terms(value: str) -> set[str]:
    return {x.casefold() for x in re.findall(r"[\w'-]+",value) if len(x)>2}


def _chunks(pages: list[PaperPage], size: int=1100, overlap: int=180) -> list[dict[str,Any]]:
    found=[]
    for p in pages:
        text=re.sub(r'\s+',' ',p.text).strip()
        start=0
        while start<len(text):
            end=min(len(text),start+size)
            cut=text.rfind('. ',start+size//2,end)
            if cut>start: end=cut+1
            found.append({'page':p.page,'text':text[start:end].strip()})
            if end==len(text): break
            start=max(start+1,end-overlap)
    return found


def paper_evidence(pages: list[PaperPage], question: str, limit: int=6) -> list[dict[str,Any]]:
    terms=_terms(question)
    if not terms: return []
    scored=[]
    for item in _chunks(pages):
        words=_terms(item['text']); overlap=terms & words
        score=len(overlap)/max(1,len(terms))
        if score: scored.append((score,item))
    scored.sort(key=lambda x:(x[0],len(x[1]['text'])),reverse=True)
    output=[]; seen=set()
    for score,item in scored:
        signature=(item['page'],item['text'][:100])
        if signature in seen: continue
        output.append({**item,'match':round(score,2)}); seen.add(signature)
        if len(output)>=limit: break
    return output


def summarize_paper(pages: list[PaperPage]) -> dict[str,Any]:
    sections=[]
    for p in pages:
        for para in re.split(r'\n\s*\n',p.text):
            clean=re.sub(r'\s+',' ',para).strip()
            if len(clean)>80: sections.append((p.page,clean))
    if not sections:
        for p in pages:
            for sentence in re.split(r'(?<=[.!?])\s+',re.sub(r'\s+',' ',p.text)):
                if len(sentence)>55: sections.append((p.page,sentence.strip()))
    if not sections: return {'overview':'The PDF contains too little readable text to create a useful summary.','key_findings':[]}
    named=[]
    for label,terms in [('Abstract',['abstract']),('Conclusion',['conclusion','conclusions']),('Findings',['findings','results']),('Methods',['methodology','methods'])]:
        for page,text in sections:
            if any(re.search(rf'\b{term}\b',text,re.I) for term in terms):
                if len(text)>180: named.append({'section':label,'page':page,'text':text[:900]}); break
    ranked=sorted(sections,key=lambda x:(len(x[1])>150, bool(re.search(r'\d|significant|increase|decrease|found|showed|demonstrate|result',x[1],re.I))),reverse=True)
    for page,text in ranked:
        if len(named)>=5: break
        if all(text not in item['text'] for item in named): named.append({'section':'Key passage','page':page,'text':text[:700]})
    return {'overview':'Extractive summary based on labeled sections and text passages. Review the cited pages for context.','key_findings':named[:5]}


def read_workbook(content: bytes) -> dict[str,pd.DataFrame]:
    try: book=pd.ExcelFile(io.BytesIO(content))
    except Exception as exc: raise ValueError('Could not open this workbook. Use a valid .xlsx or .xls file.') from exc
    result={}
    for sheet in book.sheet_names[:30]:
        frame=pd.read_excel(book,sheet_name=sheet,header=None)
        if frame.empty: continue
        header=frame.iloc[0].fillna('').astype(str).str.strip().tolist()
        if not any(header): continue
        header=[name if name else f'column_{i+1}' for i,name in enumerate(header)]
        frame=frame.iloc[1:].copy(); frame.columns=header
        frame=frame.dropna(how='all')
        result[sheet]=frame
    if not result: raise ValueError('No non-empty worksheets with a header row were found.')
    return result


def _resolve_column(frame: pd.DataFrame, query: str) -> str | None:
    aliases={'sales':['sales','revenue','amount','total','value'],'month':['month','date','period'],'product':['product','item','sku'],
             'quantity':['quantity','qty','units','volume'],'region':['region','area','territory'],'customer':['customer','client','account']}
    q=query.casefold()
    for concept,names in aliases.items():
        if any(re.search(rf'\b{re.escape(word)}\b',q) for word in names):
            for col in frame.columns:
                if any(re.search(rf'\b{re.escape(word)}\b',str(col).casefold()) for word in names): return str(col)
    for col in frame.columns:
        words=re.findall(r"[\w'-]+",str(col).casefold())
        if words and all(re.search(rf'\b{re.escape(word)}\b',q) for word in words): return str(col)
    return None


def analyze_table(frame: pd.DataFrame, question: str, row_limit: int=12) -> dict[str,Any]:
    if frame.empty: return {'answer':'This worksheet has no data rows.','evidence':[],'calculation':'No rows to analyze.'}
    q=question.casefold(); measure=_resolve_column(frame,q)
    month_col=_resolve_column(frame,'month')
    if any(x in q for x in ['highest','maximum','max','most','top']) and measure:
        numeric=pd.to_numeric(frame[measure].astype(str).str.replace(r'[^\d.-]','',regex=True),errors='coerce')
        valid=numeric.notna()
        if not valid.any(): return {'answer':f"Column '{measure}' has no values that can be read as numbers.",'evidence':[],'calculation':f'Converted {measure} values to numbers; no valid numeric cells were found.'}
        max_value=float(numeric[valid].max()); matches=frame.loc[valid & numeric.eq(max_value)].copy(); matches['_source_row']=matches.index+1
        group_col=month_col if month_col and month_col!=measure else None
        if group_col and any(word in q for word in ['month','monthly']) and ('sales' in q or 'revenue' in q or measure.casefold() in q):
            groups=frame.copy(); groups['_numeric']=numeric; groups=groups.dropna(subset=['_numeric']).groupby(group_col,dropna=False)['_numeric'].sum().sort_values(ascending=False)
            if len(groups):
                top_value=float(groups.iloc[0]); top_month=str(groups.index[0]); selected=frame[frame[group_col].astype(str)==top_month].copy(); selected['_source_row']=selected.index+1
                return {'answer':f"{top_month} has the highest total {measure}: {top_value:,.2f}.",'evidence':selected.head(row_limit).to_dict('records'),'calculation':f"Converted '{measure}' to numbers and summed it by '{group_col}'. The maximum grouped sum is {top_value:,.2f}.",'supporting_rows':len(selected),'columns':list(frame.columns)}
        label=str(matches.iloc[0][group_col]) if group_col else str(matches.iloc[0].get('_source_row',''))
        return {'answer':f"The highest {measure} value is {max_value:,.2f}"+(f' for {label}.' if group_col else '.'),'evidence':matches.head(row_limit).to_dict('records'),'calculation':f"Converted '{measure}' to numbers and selected the maximum value ({max_value:,.2f}). Rows shown use Excel source row numbers.",'supporting_rows':len(matches),'columns':list(frame.columns)}
    if any(x in q for x in ['average','mean']) and measure:
        numeric=pd.to_numeric(frame[measure].astype(str).str.replace(r'[^\d.-]','',regex=True),errors='coerce'); count=int(numeric.notna().sum())
        if count: return {'answer':f"The average of {measure} is {numeric.mean():,.2f} across {count} numeric rows.",'evidence':[],'calculation':f"Mean = sum of {count} numeric values divided by {count}.",'supporting_rows':count,'columns':list(frame.columns)}
    if any(x in q for x in ['sum','total','overall']) and measure:
        numeric=pd.to_numeric(frame[measure].astype(str).str.replace(r'[^\d.-]','',regex=True),errors='coerce'); count=int(numeric.notna().sum())
        if count: return {'answer':f"The total {measure} is {numeric.sum():,.2f} across {count} numeric rows.",'evidence':[],'calculation':f"Sum of {count} numeric values in '{measure}'.",'supporting_rows':count,'columns':list(frame.columns)}
    terms=_terms(question); scored=[]
    for _,row in frame.iterrows():
        text=' '.join(str(v) for v in row.tolist()); score=len(terms & _terms(text))
        if score:
            item=row.to_dict(); item['_source_row']=int(row.name)+1; scored.append((score,item))
    scored.sort(key=lambda x:x[0],reverse=True)
    if not scored:
        return {'answer':'I could not find a relevant value or matching row in this worksheet. Try naming a column, measure, or category.','evidence':[],'calculation':'No matching data found.','columns':list(frame.columns)}
    evidence=[item for _,item in scored[:row_limit]]
    return {'answer':f"Found {len(scored)} row(s) with words matching your question. Review the supporting rows before drawing a conclusion.",'evidence':evidence,'calculation':'Rows ranked by simple word overlap with the question; this is a search result, not a semantic or statistical conclusion.','supporting_rows':len(scored),'columns':list(frame.columns)}
