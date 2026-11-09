from __future__ import annotations
import hashlib, io, json, os
from pathlib import Path
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from document_agent.extraction import classify_document, extract_fields, extract_text, review_flags
from document_agent.providers import configured, extract_with_provider
from document_agent.storage import connect, create_job, events, jobs, review_job
from document_agent.ask_files import analyze_table, paper_evidence, read_paper_pdf, read_workbook, summarize_paper

ROOT=Path(__file__).resolve().parent
DATA=ROOT/'data'; DATA.mkdir(exist_ok=True); load_dotenv(ROOT/'.env')
st.set_page_config(page_title='DocuMind · Document intelligence',page_icon=':material/description:',layout='wide')

@st.cache_resource
def get_db(): return connect(os.getenv('DOCINTEL_DB_PATH',str(DATA/'document_intelligence.sqlite')))

def rows_now(): return jobs(get_db())
def decoded(row): return json.loads(row['extracted_json'])
def flags(row): return json.loads(row['flags_json'])
def record(row): return {'job_id':row['id'],'filename':row['filename'],'type':row['doc_type'],'status':row['status'],'created_at':row['created_at'],**decoded(row)}
def workbook(frame):
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as writer: frame.to_excel(writer,index=False,sheet_name='Records')
    return out.getvalue()
def review_workbook(selected_rows):
    """Build a human-editable Excel review queue, with instructions and flat field columns."""
    from openpyxl.styles import Font, PatternFill
    out=io.BytesIO()
    payload=[]
    for row in selected_rows:
        item={'job_id':row['id'],'filename':row['filename'],'document_type':row['doc_type'],
              'review_decision':row['status'],'reviewer_note':row['reviewer_note']}
        for key,value in decoded(row).items():
            item[f'field__{key}']=', '.join(map(str,value)) if isinstance(value,list) else value
        payload.append(item)
    with pd.ExcelWriter(out,engine='openpyxl') as writer:
        pd.DataFrame([{'Step':'1. Edit field values in the Review sheet. Keep job_id unchanged.',
                       'Decision values':'pending, approved, needs_changes',
                       'Privacy':'This workbook contains extracted document data. Store and share it carefully.'},
                      {'Step':'2. Save the workbook as .xlsx, then import it in DocuMind → Records & exports.',
                       'Decision values':'Do not rename or remove job_id, review_decision, or reviewer_note.',
                       'Privacy':'Import updates existing records; it does not add new documents.'},
                      {'Step':'3. Confirm the import preview before applying updates.',
                       'Decision values':'Imported edits and decisions are written to the SQLite audit trail.',
                       'Privacy':'Verify every value against the original source before approval.'}]).to_excel(writer,index=False,sheet_name='Start here')
        pd.DataFrame(payload).to_excel(writer,index=False,sheet_name='Review')
        pd.DataFrame([record(row) for row in selected_rows]).to_excel(writer,index=False,sheet_name='Snapshot')
        for sheet in writer.book.worksheets:
            sheet.freeze_panes='A2'; sheet.auto_filter.ref=sheet.dimensions
            for cell in sheet[1]:
                cell.font=Font(bold=True,color='FFFFFF')
                cell.fill=PatternFill('solid',fgColor='256B58')
            for column in sheet.columns:
                width=min(max(max(len(str(cell.value or '')) for cell in list(column)[:200])+2,12),42)
                sheet.column_dimensions[column[0].column_letter].width=width
    return out.getvalue()

def validate_review_sheet(upload, current_rows):
    frame=pd.read_excel(upload,sheet_name='Review',dtype={'job_id':'Int64'})
    required={'job_id','review_decision','reviewer_note'}
    if not required.issubset(frame.columns): raise ValueError('The Review sheet must contain job_id, review_decision, and reviewer_note columns.')
    if frame['job_id'].isna().any() or frame['job_id'].duplicated().any(): raise ValueError('Each row needs a unique, non-empty job_id.')
    valid={int(row['id']):row for row in current_rows}
    updates=[]
    allowed={'pending','approved','needs_changes'}
    for _,item in frame.iterrows():
        job_id=int(item['job_id'])
        if job_id not in valid: raise ValueError(f'Unknown job_id {job_id}; this workbook may belong to a different lab database.')
        decision=str(item['review_decision']).strip().lower()
        if decision not in allowed: raise ValueError(f'Invalid review_decision for job {job_id}. Use pending, approved, or needs_changes.')
        before=decoded(valid[job_id]); updated={}
        for key in before:
            column=f'field__{key}'
            value=item[column] if column in frame.columns else before[key]
            if pd.isna(value): value=None
            if isinstance(before[key],list):
                updated[key]=[part.strip() for part in str(value).split(',') if part.strip()] if value is not None else []
            else: updated[key]=None if value is None or str(value).strip()=='' else str(value).strip()
        note=item['reviewer_note']
        updates.append((job_id,updated,decision,'' if pd.isna(note) else str(note)))
    return updates
def hero(kicker,title,subtitle):
    st.caption(kicker.upper())
    st.title(title)
    st.caption(subtitle)
def pill_label(value): return {'approved':'Approved','pending':'Needs review','needs_changes':'Changes requested'}.get(value,value.replace('_',' ').title())

db=get_db()
with st.sidebar:
    st.markdown('## :material/auto_awesome: DocuMind')
    st.caption('DOCUMENT INTELLIGENCE STUDIO')
    page=st.radio('Workspace',['Overview','Ask your files','Ingest documents','Review desk','Records & exports','Audit trail'],label_visibility='collapsed',format_func=lambda x:{'Overview':':material/dashboard:  Overview','Ask your files':':material/manage_search:  Ask your files','Ingest documents':':material/cloud_upload:  Ingest documents','Review desk':':material/fact_check:  Review desk','Records & exports':':material/database:  Records & exports','Audit trail':':material/history:  Audit trail'}[x],key='workspace_page')
    st.divider()
    all_rows=rows_now(); pending=sum(r['status']!='approved' for r in all_rows)
    st.metric('Awaiting review',pending)
    st.caption('Local processing · English + Hindi')
    st.caption('Uploads are not retained as files.')

rows=rows_now()
approved=sum(r['status']=='approved' for r in rows)
pending=sum(r['status']!='approved' for r in rows)
needs=sum(r['status']=='needs_changes' for r in rows)

if page=='Overview':
    hero('Document intelligence studio','Your documents, in focus.','Turn PDFs and scans into reviewable records, with every decision kept traceable.')
    st.write('')
    with st.container(border=True):
        st.markdown('**Workspace snapshot**')
        c1,c2,c3,c4=st.columns(4)
        c1.metric('Documents processed',len(rows))
        c2.metric('Awaiting review',pending)
        c3.metric('Approved records',approved)
        c4.metric('Changes requested',needs)
    st.write('')
    left,right=st.columns([1.35,1],gap='large')
    with left:
        with st.container(border=True):
            st.markdown('### Recent documents')
            st.caption('Latest files added to this workspace')
            if rows:
                for row in rows[:6]:
                    with st.container(horizontal=True,vertical_alignment='center'):
                        st.markdown(f"**{row['filename']}**  ·  {row['doc_type'].title()}")
                        st.badge(pill_label(row['status']),color='green' if row['status']=='approved' else 'orange' if row['status']=='pending' else 'red')
                        st.caption(row['created_at'][:16].replace('T',' '))
            else:
                st.markdown('No documents yet. Start by adding a PDF or image to create your first review draft.')
                if st.button('Go to document intake',icon=':material/upload_file:',key='empty_ingest'):
                    st.session_state['workspace_page']='Ingest documents'; st.rerun()
    with right:
        with st.container(border=True):
            st.markdown('### A clear review path')
            st.markdown('**01 · Ingest**  Add PDFs or images and choose OCR languages.')
            st.markdown('**02 · Inspect**  Review extracted values and source text side by side.')
            st.markdown('**03 · Decide**  Edit the record and approve or request changes.')
            st.markdown('**04 · Export**  Download the records you need in CSV, Excel, or JSON.')
            st.markdown('**05 · Work in Excel**  Download the review workbook, edit it, then import your reviewed rows.')
        with st.container(border=True):
            st.markdown('### Privacy at a glance')
            st.caption('Text extraction and OCR run locally. Uploaded files are held in memory during processing; SQLite saves structured fields and an integrity digest. No external provider is used unless configured and explicitly authorized.')
            st.caption('No login, encryption-at-rest management, or retention controls are included in this MVP.')
    with st.container(border=True):
        st.markdown('### Use Excel with DocuMind')
        st.markdown('1. Open **Records & exports** and download the **Excel review workbook**.\n2. In Excel, edit columns starting with `field__` and set `review_decision` to `pending`, `approved`, or `needs_changes`. Add any reviewer explanation in `reviewer_note`.\n3. Keep `job_id` unchanged. Save as `.xlsx`.\n4. Upload the file on the same **Records & exports** page, review the preview, tick the confirmation box, and apply the updates.\n5. Check **Audit trail** to see the imported edits and decisions.')
        st.caption('The workbook includes Start here, Review, and Snapshot sheets. Only the Review sheet is used for imports. Verify all values against the original documents before approving.')

elif page=='Ask your files':
    hero('Explore and question','Ask your files','Summarize research PDFs, search for findings, and analyze spreadsheet data with evidence you can inspect.')
    st.info('Analysis runs locally in this app. Uploaded files are kept in memory for this browser session and are not added to the SQLite document register. Answers use extracted text, simple retrieval, and deterministic calculations—not an external LLM.',icon=':material/lock:')
    mode=st.segmented_control('Choose a file type',['Research paper (PDF)','Spreadsheet (Excel)'],default='Research paper (PDF)',key='ask_mode')
    if mode=='Research paper (PDF)':
        with st.container(border=True):
            paper=st.file_uploader('Upload a research paper PDF',type=['pdf'],key='ask_paper_pdf',help='Up to 20 MB and 200 pages. Scanned pages use local Tesseract OCR.')
            paper_language=st.selectbox('OCR language for scanned pages',['English + Hindi','English only','Hindi only'],key='ask_paper_language')
            paper_lang_code={'English + Hindi':'eng+hin','English only':'eng','Hindi only':'hin'}[paper_language]
            load_paper=st.button('Read research paper',type='primary',icon=':material/auto_stories:',disabled=not paper)
        if load_paper and paper:
            try:
                content=paper.getvalue()
                if len(content)>20*1024*1024: raise ValueError('File exceeds the 20 MB limit.')
                parsed=read_paper_pdf(paper.name,content,paper_lang_code)
                st.session_state['ask_paper_pages']=[{'page':p.page,'text':p.text} for p in parsed]
                st.session_state['ask_paper_name']=paper.name
                st.session_state.pop('ask_paper_summary',None)
                st.success(f"Read {len(parsed)} pages from {paper.name}. Text stays in this session.")
            except Exception as exc: st.error(str(exc))
        stored=st.session_state.get('ask_paper_pages')
        if stored:
            parsed=[type('Page',(),item) for item in stored]
            st.caption(f"Current paper: **{st.session_state.get('ask_paper_name','Uploaded PDF')}** · {len(parsed)} pages")
            summary_col,question_col=st.columns([1,1],gap='large')
            with summary_col:
                with st.container(border=True):
                    st.markdown('### Summarize this paper')
                    st.caption('Extractive summary: key passages from labels such as abstract, methods, results, and conclusion, with page numbers.')
                    if st.button('Create paper summary',icon=':material/summarize:'):
                        st.session_state['ask_paper_summary']=summarize_paper(parsed)
                    summary=st.session_state.get('ask_paper_summary')
                    if summary:
                        st.write(summary['overview'])
                        for item in summary['key_findings']:
                            with st.expander(f"{item['section']} · page {item['page']}"):
                                st.write(item['text'])
            with question_col:
                with st.container(border=True):
                    st.markdown('### Find an answer')
                    question=st.text_area('Ask about this paper',placeholder='What were the main findings? Which method did the authors use?',height=95,key='paper_question')
                    if st.button('Search paper evidence',type='primary',icon=':material/search:'):
                        hits=paper_evidence(parsed,question)
                        st.session_state['ask_paper_hits']={'question':question,'hits':hits}
                    result=st.session_state.get('ask_paper_hits')
                    if result:
                        if result['hits']:
                            st.markdown(f"**Relevant passages for:** {result['question']}")
                            for item in result['hits']:
                                with st.expander(f"Page {item['page']} · term match {item['match']:.0%}"):
                                    st.write(item['text'])
                            st.caption('Passages are selected by word overlap. Read the cited page and surrounding context; relevance is not proof that the passage fully answers the question.')
                        else: st.warning('No matching evidence found in the paper text. Try wording the question with terms used by the authors.')
            with st.expander('Browse extracted page text'):
                page_choice=st.selectbox('Page',[p.page for p in parsed],key='paper_page_choice')
                st.text_area('Extracted text',next(p.text for p in parsed if p.page==page_choice),height=300,key=f'paper_text_{page_choice}')
    else:
        with st.container(border=True):
            excel=st.file_uploader('Upload an Excel workbook',type=['xlsx'],key='ask_excel',help='Up to 20 MB. The first row in each worksheet is treated as the column header.')
            load_excel=st.button('Read spreadsheet',type='primary',icon=':material/table_view:',disabled=not excel)
            sample_excel=st.button('Try a synthetic sales example',icon=':material/science:')
        if sample_excel:
            example=pd.DataFrame({'Month':['January','January','February','February','March','March','April','April'],
                                  'Sales':[12500,9800,14200,15600,18100,17300,16400,19300],
                                  'Product':['Notebook','Folders','Notebook','Folders','Notebook','Folders','Notebook','Folders'],
                                  'Region':['North','South','North','South','North','South','North','South']},index=range(1,9))
            st.session_state['ask_excel_sheets']={'Synthetic sales sample':example.to_json(orient='split')}
            st.session_state['ask_excel_name']='Synthetic sales sample (fictional)'
            st.session_state.pop('ask_excel_result',None)
            st.success('Loaded a small fictional dataset. Try “Which month had the highest sales?”')
        if load_excel and excel:
            try:
                content=excel.getvalue()
                if len(content)>20*1024*1024: raise ValueError('File exceeds the 20 MB limit.')
                sheets=read_workbook(content)
                st.session_state['ask_excel_sheets']={name:frame.to_json(orient='split',date_format='iso') for name,frame in sheets.items()}
                st.session_state['ask_excel_name']=excel.name
                st.session_state.pop('ask_excel_result',None)
                st.success(f"Read {len(sheets)} worksheets from {excel.name}. Workbook data stays in this session.")
            except Exception as exc: st.error(str(exc))
        stored_sheets=st.session_state.get('ask_excel_sheets')
        if stored_sheets:
            st.caption(f"Current workbook: **{st.session_state.get('ask_excel_name','Uploaded workbook')}** · {len(stored_sheets)} sheets")
            sheet_name=st.selectbox('Choose a worksheet',list(stored_sheets),key='ask_excel_sheet_choice')
            frame=pd.read_json(io.StringIO(stored_sheets[sheet_name]),orient='split')
            with st.container(border=True):
                st.markdown('### Spreadsheet preview')
                st.caption(f'{len(frame):,} rows · {len(frame.columns)} columns · source row numbers include the header row')
                st.dataframe(frame.head(12),hide_index=True,width='stretch')
            with st.container(border=True):
                st.markdown('### Ask about this data')
                question=st.text_input('Your question',placeholder='Which month had the highest sales?',key='excel_question')
                if st.button('Analyze worksheet',type='primary',icon=':material/query_stats:',disabled=not question):
                    st.session_state['ask_excel_result']=analyze_table(frame,question)
                result=st.session_state.get('ask_excel_result')
                if result:
                    st.markdown('#### Answer')
                    st.write(result['answer'])
                    st.caption(result['calculation'])
                    if result.get('evidence'):
                        st.markdown('**Supporting rows**')
                        st.dataframe(pd.DataFrame(result['evidence']),hide_index=True,width='stretch')
                    elif result.get('supporting_rows'):
                        st.caption(f"Computed from {result['supporting_rows']} matching numeric row(s).")
                    with st.expander('Columns and data used'):
                        st.write(result.get('columns',list(frame.columns)))
                        st.caption('Numeric text is parsed by removing currency symbols and separators. Check source cells, units, blanks, and grouping before relying on the result.')
    st.write('')
    with st.container(border=True):
        st.markdown('### What is in the queue?')
        if rows:
            dist=pd.DataFrame([{'Document type':r['doc_type'].title(),'Count':1} for r in rows]).groupby('Document type',as_index=False)['Count'].sum()
            st.bar_chart(dist.set_index('Document type'),height=220,color='#256B58')
        else: st.caption('Your document type summary will appear here after processing files.')

elif page=='Ingest documents':
    hero('01 / Intake','Bring documents into the workspace','Files are validated and processed on this machine. Only extracted records and audit details are saved.')
    a,b=st.columns([1.15,.85],gap='large')
    with a:
        with st.container(border=True):
            st.markdown('### Upload documents')
            uploads=st.file_uploader('Choose PDFs or images',type=['pdf','png','jpg','jpeg','tif','tiff','bmp'],accept_multiple_files=True,help='Up to 20 MB each. PDFs are limited to 60 pages.')
            st.caption('Supported: PDF, PNG, JPG, TIFF, BMP · up to 20 MB per file')
            c1,c2=st.columns(2)
            languages=c1.selectbox('OCR languages',['English + Hindi','English only','Hindi only'])
            lang_code={'English + Hindi':'eng+hin','English only':'eng','Hindi only':'hin'}[languages]
            doc_hint=c2.selectbox('Document types',['Auto-detect','Resume','Certificate','Invoice','Form','Other'])
            custom=st.text_input('Custom fields',placeholder='e.g. policy number, department',help='Comma-separated labels to search for in extracted text.')
            process=st.button('Extract records',type='primary',icon=':material/auto_awesome:',disabled=not uploads,width='stretch')
    with b:
        with st.container(border=True):
            st.markdown('### Processing pipeline')
            st.markdown('1. **Validate** file type, signature, size and page count')
            st.markdown('2. **Read** embedded PDF text or run local OCR for scans')
            st.markdown('3. **Structure** classify and suggest field values')
            st.markdown('4. **Review** save a draft for a person to check')
            st.caption('Classification and field suggestions use transparent starter rules. They can miss values and need source verification.')
        with st.container(border=True):
            st.markdown('### OCR readiness')
            st.markdown('**English + Hindi**')
            st.caption('Requires local Tesseract OCR and the selected language data. Text-based PDFs can be read without Tesseract.')
            st.caption('Setup guide and product limitations are documented in README.md.')
    if process:
        st.divider(); st.markdown('### Processing results')
        progress=st.progress(0,text='Preparing files…'); result_area=st.container()
        for index,up in enumerate(uploads):
            try:
                content=up.getvalue(); result=extract_text(up.name,content,lang_code); detected,_=classify_document(result.text)
                chosen=detected if doc_hint=='Auto-detect' else doc_hint.casefold()
                fields=extract_fields(result.text,chosen,[x.strip() for x in custom.split(',') if x.strip()])
                warnings=review_flags(fields,chosen,result.text)+[{'level':'review','message':w} for w in result.warnings]
                job_id=create_job(db,filename=up.name,sha256=hashlib.sha256(content).hexdigest(),doc_type=chosen,language=lang_code,method=result.method,fields=fields,flags=warnings)
                st.session_state[f'text_{job_id}']=result.text
                with result_area:
                    with st.container(border=True):
                        st.markdown(f"**{up.name}** · {chosen.title()} · {result.method}")
                        st.success(f'Review draft #{job_id} created. {len(warnings)} item(s) to check.')
                        for item in warnings: st.warning(item['message'])
                        with st.expander('Preview extracted text'):
                            st.text_area('Text preview',result.text[:5000],height=180,key=f'ingest_preview_{job_id}')
            except Exception as exc:
                with result_area: st.error(f'{up.name}: {exc}',icon=':material/error:')
            progress.progress((index+1)/len(uploads),text=f'Processed {index+1} of {len(uploads)} files')
        if uploads and st.button('Open review desk',icon=':material/fact_check:'):
            st.session_state['workspace_page']='Review desk'; st.rerun()
    st.divider(); st.markdown('### Optional provider suggestions')
    if configured():
        with st.container(border=True):
            st.warning('Provider calls transmit extracted document text outside this machine.',icon=':material/cloud:')
            consent=st.checkbox('I have reviewed provider data handling and authorize sending a selected document text for suggestions.')
            current=rows_now(); available=[r for r in current if st.session_state.get(f"text_{r['id']}")]
            if available:
                selected=st.selectbox('Select an in-session document',[r['id'] for r in available],format_func=lambda i:next(f"#{r['id']} · {r['filename']}" for r in available if r['id']==i))
                if st.button('Request field suggestions',disabled=not consent,type='secondary'):
                    row=next(r for r in available if r['id']==selected)
                    try:
                        suggestion=extract_with_provider(st.session_state[f'text_{selected}'],row['doc_type'],decoded(row))
                        st.json(suggestion); st.caption('Provider response is unverified and is not automatically saved or approved.')
                    except Exception as exc: st.error(f'Provider request failed safely: {exc}')
            else: st.caption('Process a document in this session before requesting provider suggestions.')
    else: st.caption('No provider configured. Local deterministic extraction remains available.')

elif page=='Review desk':
    hero('02 / Human review','Review desk','Check extracted values against the source text, make corrections, and record a decision.')
    if not rows: st.info('No records are waiting yet. Add documents to open a review draft.',icon=':material/inbox:')
    else:
        f1,f2,f3=st.columns([1,1,1.4],vertical_alignment='bottom')
        status_filter=f1.segmented_control('Review state',['All','Needs review','Approved','Changes requested'],default='Needs review',key='review_status')
        type_filter=f2.selectbox('Document type',['All']+sorted({r['doc_type'].title() for r in rows}))
        query=f3.text_input('Find a document',placeholder='Search filename or extracted text',icon=':material/search:')
        filtered=[]
        for row in rows:
            state={'Approved':'approved','Changes requested':'needs_changes'}.get(status_filter)
            if status_filter=='Needs review' and row['status']=='approved': continue
            if state and row['status']!=state: continue
            if type_filter!='All' and row['doc_type'].title()!=type_filter: continue
            search_blob=(row['filename']+' '+json.dumps(decoded(row),ensure_ascii=False)).casefold()
            if query and query.casefold() not in search_blob: continue
            filtered.append(row)
        if not filtered: st.info('No records match these filters.',icon=':material/filter_list_off:')
        else:
            pick=st.selectbox('Select a record to inspect',[r['id'] for r in filtered],format_func=lambda i:next(f"{r['filename']}  ·  {r['doc_type'].title()}  ·  {pill_label(r['status'])}" for r in filtered if r['id']==i),label_visibility='collapsed')
            row=next(r for r in filtered if r['id']==pick); fields=decoded(row)
            main,source=st.columns([1.1,.9],gap='large')
            with main:
                with st.container(border=True):
                    st.markdown(f"### {row['filename']}")
                    st.badge(pill_label(row['status']),color='green' if row['status']=='approved' else 'orange' if row['status']=='pending' else 'red')
                    st.caption(f"{row['doc_type'].title()} · {row['method']} · added {row['created_at'][:16].replace('T',' ')}")
                    for item in flags(row): st.warning(item['message'])
                    st.markdown('#### Suggested fields')
                    edited={}
                    with st.form(f'review_form_{pick}'):
                        for key,value in fields.items():
                            if isinstance(value,list):
                                raw=st.text_input(key.replace('_',' ').title(),value=', '.join(map(str,value)),help='Comma-separated values')
                                edited[key]=[v.strip() for v in raw.split(',') if v.strip()]
                            else:
                                edited[key]=st.text_input(key.replace('_',' ').title(),value='' if value is None else str(value)) or None
                        note=st.text_area('Reviewer note',value=row['reviewer_note'],placeholder='What did you verify or correct?')
                        decision=st.segmented_control('Decision',['approved','needs_changes'],default=row['status'] if row['status'] in ('approved','needs_changes') else 'approved',format_func=lambda x:'Approve' if x=='approved' else 'Request changes')
                        submit=st.form_submit_button('Save review decision',type='primary',icon=':material/save:')
                    if submit:
                        review_job(db,pick,edited,decision or 'approved',note)
                        st.success('Decision and field edits saved to the audit trail.')
                        st.rerun()
            with source:
                with st.container(border=True):
                    st.markdown('### Source context')
                    extracted_text=st.session_state.get(f'text_{pick}')
                    if extracted_text:
                        st.caption('In-memory text from the current session. Source files are not stored.')
                        st.text_area('Extracted source text',extracted_text,height=440,key=f'source_{pick}')
                    else:
                        st.info('Source text is not retained between sessions. Re-upload the document to inspect it again.',icon=':material/description:')
                        st.caption(f"SHA-256: `{row['sha256']}`")
                    st.download_button('Download record JSON',json.dumps(record(row),ensure_ascii=False,indent=2),file_name=f'document-{pick}.json',mime='application/json',icon=':material/download:')

elif page=='Records & exports':
    hero('03 / Workspace data','Records & exports','Filter the saved register and export exactly the records you want to share.')
    if not rows: st.info('No saved records to display.',icon=':material/database:')
    else:
        f1,f2,f3=st.columns([1,1,1.5],vertical_alignment='bottom')
        type_filter=f1.multiselect('Document type',sorted({r['doc_type'].title() for r in rows}),placeholder='All types')
        status_filter=f2.multiselect('Review state',['pending','approved','needs_changes'],placeholder='All states')
        query=f3.text_input('Search records',placeholder='Filename or field value',icon=':material/search:')
        filtered=[]
        for row in rows:
            if type_filter and row['doc_type'].title() not in type_filter: continue
            if status_filter and row['status'] not in status_filter: continue
            if query and query.casefold() not in (row['filename']+' '+json.dumps(decoded(row),ensure_ascii=False)).casefold(): continue
            filtered.append(row)
        data=[record(r) for r in filtered]; frame=pd.json_normalize(data,sep='.')
        st.caption(f'{len(filtered)} of {len(rows)} records selected')
        if data:
            x1,x2,x3=st.columns(3)
            x1.download_button('Export CSV',frame.to_csv(index=False).encode('utf-8-sig'),'document_records.csv','text/csv',icon=':material/table_view:')
            x2.download_button('Export Excel snapshot',workbook(frame),'document_records.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',icon=':material/table_chart:')
            x3.download_button('Export JSON',json.dumps(data,ensure_ascii=False,indent=2),'document_records.json','application/json',icon=':material/data_object:')
            st.download_button('Download Excel review workbook',review_workbook(filtered),'document_review_workbook.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',icon=':material/edit_document:')
            st.dataframe(pd.DataFrame([{'ID':r['id'],'Document':r['filename'],'Type':r['doc_type'].title(),'Review state':pill_label(r['status']),'OCR / text method':r['method'],'Created':r['created_at'],'Reviewed':r['reviewed_at'] or '—'} for r in filtered]),hide_index=True,width='stretch',selection_mode='single-row',on_select='rerun',key='records_table')
            selected=st.session_state.records_table.selection.rows
            if selected:
                chosen=filtered[selected[0]]
                with st.container(border=True):
                    st.markdown(f"### {chosen['filename']}")
                    st.json(decoded(chosen))
        st.divider()
        st.markdown('### Bring reviewed rows back from Excel')
        st.caption('Upload a completed copy of the Excel review workbook. The lab previews every row and validates IDs and decisions before you choose to apply changes.')
        excel_upload=st.file_uploader('Upload completed review workbook',type=['xlsx'],key='excel_review_upload')
        if excel_upload:
            try:
                updates=validate_review_sheet(excel_upload,rows)
                st.success(f'{len(updates)} existing record(s) passed validation.')
                preview=[]
                for job_id,values,decision,note in updates:
                    source=next(r for r in rows if r['id']==job_id)
                    preview.append({'job_id':job_id,'filename':source['filename'],'decision':decision,'fields':json.dumps(values,ensure_ascii=False),'reviewer_note':note})
                st.dataframe(pd.DataFrame(preview),hide_index=True,width='stretch')
                confirm=st.checkbox('I checked these edits against the source documents and want to apply them to the lab database.',key='excel_import_confirm')
                if st.button('Apply Excel review updates',type='primary',disabled=not confirm,icon=':material/save:'):
                    for job_id,values,decision,note in updates: review_job(db,job_id,values,decision,note)
                    st.success('Excel edits imported. Each record update is now in the audit trail.')
                    st.rerun()
            except Exception as exc:
                st.error(f'Could not import this workbook: {exc}')
        else: st.info('No records match these filters.')

elif page=='Audit trail':
    hero('04 / Traceability','Audit trail','A chronological record of drafts, field edits, and human decisions.')
    selected=st.selectbox('Filter by document',['All documents']+[f"#{r['id']} · {r['filename']}" for r in rows])
    job_id=None if selected=='All documents' else int(selected.split(' · ',1)[0][1:])
    history=events(db,job_id)
    if history:
        for event in history:
            details=json.loads(event['details_json'])
            title=f"{event['event'].replace('_',' ').title()} · document #{event['job_id']}"
            with st.expander(f"{event['created_at']}  ·  {title}"):
                st.json(details)
    else: st.info('No audit events yet.',icon=':material/history:')
    st.caption('Audit history records user decisions and the previous/current review state. It does not certify document authenticity.')
