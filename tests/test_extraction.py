import io
import fitz
import pytest
from PIL import Image
from document_agent.extraction import classify_document, extract_fields, review_flags, validate_upload

def test_upload_type_size_and_signature():
    with pytest.raises(ValueError): validate_upload('x.exe',b'hello')
    with pytest.raises(ValueError): validate_upload('x.pdf',b'not pdf')
    with pytest.raises(ValueError): validate_upload('x.png',b'')
    with pytest.raises(ValueError): validate_upload('x.png',b'x'*(20*1024*1024+1))
    buf=io.BytesIO(); Image.new('RGB',(4,4),'white').save(buf,format='PNG'); validate_upload('image.png',buf.getvalue())

def test_classifier_and_invoice_extraction():
    text='Invoice No: INV-42\nVendor: North Star\nGrand Total: INR 1,250.00\nName: Asha Rao\nDate: 2026-03-01'
    kind,_=classify_document(text); fields=extract_fields(text,kind)
    assert kind=='invoice'
    assert fields['name']=='Asha Rao'
    assert fields['invoice_number']=='INV-42'
    assert fields['total']=='INR 1,250.00'
    assert fields['dates']==['2026-03-01']

def test_hindi_patterns_and_review_flags_are_cautious():
    text='प्रमाणपत्र\nनाम: सीमा देवी\nप्रमाणपत्र में डेटा विज्ञान'
    kind,_=classify_document(text); fields=extract_fields(text,kind)
    assert kind=='certificate' and fields['name']=='सीमा देवी'
    assert all('fraud' not in f['message'].lower() for f in review_flags(fields,kind,text))

def test_pdf_text_extraction_without_ocr():
    from document_agent.extraction import extract_text
    doc=fitz.open(); page=doc.new_page(); page.insert_text((72,72),'Invoice No: INV-9\nGrand Total: INR 99.00')
    raw=doc.tobytes(); doc.close()
    result=extract_text('sample.pdf',raw)
    assert 'INV-9' in result.text and result.method=='PDF text'
