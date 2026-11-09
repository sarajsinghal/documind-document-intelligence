import json
from document_agent.storage import connect, create_job, events, jobs, review_job

def test_job_review_and_audit_trail(tmp_path):
    db=connect(tmp_path/'test.sqlite')
    job_id=create_job(db,filename='sample.pdf',sha256='abc',doc_type='invoice',language='eng',method='PDF text',fields={'total':'10'},flags=[])
    review_job(db,job_id,{'total':'10.00'},'approved','checked against source')
    row=jobs(db)[0]
    assert row['status']=='approved' and json.loads(row['extracted_json'])['total']=='10.00'
    history=events(db,job_id)
    assert [e['event'] for e in history]==['reviewed','created']
    assert json.loads(history[0]['details_json'])['fields_changed'] is True
