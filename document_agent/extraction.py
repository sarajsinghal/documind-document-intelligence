"""Local PDF/image text, classification, and deterministic structured extraction."""
from __future__ import annotations

import io
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import fitz
from PIL import Image, ImageOps

MAX_FILE_BYTES = 20 * 1024 * 1024
ALLOWED_EXTENSIONS = {"pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp"}
DATE_RE = re.compile(r"\b(?:\d{1,2}[/-]\d{1,2}[/-](?:\d{2}|\d{4})|\d{4}-\d{2}-\d{2})\b")
AMOUNT_RE = re.compile(r"(?:₹|INR\s*|Rs\.?\s*)\s?\d[\d,]*(?:\.\d{1,2})?|\b\d[\d,]*\.\d{2}\b", re.I)


@dataclass
class ExtractionResult:
    text: str
    method: str
    warnings: list[str]


def validate_upload(name: str, content: bytes) -> None:
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if ext not in ALLOWED_EXTENSIONS:
        raise ValueError("Unsupported file type. Upload a PDF or common image file.")
    if not content:
        raise ValueError("The uploaded file is empty.")
    if len(content) > MAX_FILE_BYTES:
        raise ValueError("File exceeds the 20 MB upload limit.")
    if ext == "pdf" and not content.startswith(b"%PDF-"):
        raise ValueError("The file extension and PDF content do not match.")
    if ext != "pdf":
        try:
            Image.open(io.BytesIO(content)).verify()
        except Exception as exc:
            raise ValueError("The image could not be validated. Use a supported image file.") from exc


def _ocr(image: Image.Image, languages: str) -> str:
    try:
        import pytesseract
        if os.getenv("TESSERACT_CMD"):
            pytesseract.pytesseract.tesseract_cmd = os.environ["TESSERACT_CMD"]
        return pytesseract.image_to_string(ImageOps.grayscale(image), lang=languages, config="--psm 3")
    except Exception as exc:
        message = str(exc).lower()
        if "not installed" in message or "not found" in message or "tesseract" in message:
            raise RuntimeError("OCR is unavailable. Install Tesseract and Hindi/English language data; see README.") from exc
        raise RuntimeError("OCR failed for this page. Check local Tesseract setup and try another file.") from exc


def extract_text(name: str, content: bytes, languages: str = "eng+hin") -> ExtractionResult:
    validate_upload(name, content)
    warnings: list[str] = []
    if name.lower().endswith(".pdf"):
        try:
            doc = fitz.open(stream=content, filetype="pdf")
            if doc.page_count > 60:
                raise ValueError("PDFs are limited to 60 pages.")
            chunks, used_ocr = [], False
            for page in doc:
                txt = page.get_text("text").strip()
                if len(txt) < 30:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                    txt = _ocr(Image.open(io.BytesIO(pix.tobytes("png"))), languages)
                    used_ocr = True
                chunks.append(txt)
            doc.close()
            if not any(chunks): warnings.append("No readable text was extracted; review the source manually.")
            return ExtractionResult("\n\n".join(chunks), "PDF text + OCR" if used_ocr else "PDF text", warnings)
        except (ValueError, RuntimeError):
            raise
        except Exception as exc:
            raise ValueError("Could not read this PDF. It may be damaged or password protected.") from exc
    try:
        text = _ocr(Image.open(io.BytesIO(content)), languages)
        if not text.strip(): warnings.append("No readable text was extracted; review the source manually.")
        return ExtractionResult(text, "Tesseract OCR", warnings)
    except RuntimeError:
        raise


def classify_document(text: str) -> tuple[str, dict[str, int]]:
    t = text.casefold()
    rules = {
        "resume": ["resume", "curriculum vitae", "experience", "education", "कौशल", "अनुभव", "शिक्षा"],
        "certificate": ["certificate", "certified", "this is to certify", "प्रमाणपत्र", "प्रमाणित"],
        "invoice": ["invoice", "bill to", "invoice no", "subtotal", "total due", "चालान", "कुल राशि"],
        "form": ["application form", "date of birth", "नाम:", "आवेदन पत्र", "जन्म तिथि"],
    }
    scores = {kind: sum(t.count(term) for term in terms) for kind, terms in rules.items()}
    winner = max(scores, key=scores.get)
    return (winner if scores[winner] else "other"), scores


def _first(patterns: list[str], text: str, flags: int = re.I | re.M) -> str | None:
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return match.group(1).strip(" \t:|-")
    return None


def extract_fields(text: str, kind: str, custom_fields: list[str] | None = None) -> dict[str, Any]:
    fields: dict[str, Any] = {
        "name": _first([r"(?:full\s+name|candidate\s+name|name)\s*[:\-]\s*([^\n,;]{2,80})", r"(?:नाम|पूरा नाम)\s*[:\-]\s*([^\n,;]{2,80})"], text),
        "dates": list(dict.fromkeys(DATE_RE.findall(text)))[:20],
        "amounts": list(dict.fromkeys(AMOUNT_RE.findall(text)))[:20],
        "qualifications": [],
    }
    qualifications = ["Ph.D.", "M.Tech", "B.Tech", "MBA", "B.Sc", "M.Sc", "B.Com", "M.Com", "Bachelor", "Master", "डिप्लोमा", "स्नातक", "परास्नातक"]
    fields["qualifications"] = [q for q in qualifications if re.search(re.escape(q), text, re.I)]
    if kind == "invoice":
        fields["invoice_number"] = _first([r"(?:invoice\s*(?:no|number|#)|चालान\s*(?:संख्या|नं))\s*[:#\-]?\s*([\w/-]+)"], text)
        fields["vendor"] = _first([r"(?:vendor|supplier|from)\s*[:\-]\s*([^\n]{2,80})"], text)
        fields["total"] = _first([r"(?:grand\s+total|total\s+due|total|कुल राशि)\s*[:\s]*((?:₹\s*|INR\s*|Rs\.?\s*)?[\d,]+(?:\.\d{1,2})?)"], text)
    elif kind == "certificate":
        fields["certificate_title"] = _first([r"(?:certificate\s+(?:of|in)|प्रमाणपत्र\s*(?:का|में))\s*([^\n]{2,90})"], text)
        fields["institution"] = _first([r"(?:issued\s+by|institution|university|संस्था|विश्वविद्यालय)\s*[:\-]?\s*([^\n]{2,90})"], text)
    elif kind == "resume":
        fields["email"] = _first([r"([\w.+-]+@[\w.-]+\.[A-Za-z]{2,})"], text)
        fields["phone"] = _first([r"((?:\+?\d[\d ()-]{7,}\d))"], text)
    elif kind == "form":
        fields["date_of_birth"] = _first([r"(?:date\s+of\s+birth|dob|जन्म तिथि)\s*[:\-]?\s*([^\n]{4,30})"], text)
        fields["address"] = _first([r"(?:address|पता)\s*[:\-]\s*([^\n]{4,120})"], text)
    if custom_fields:
        for label in custom_fields:
            safe = re.escape(label)
            fields[label] = _first([rf"{safe}\s*[:\-]\s*([^\n,;]{{1,120}})"], text)
    return fields


def review_flags(fields: dict[str, Any], kind: str, text: str) -> list[dict[str, str]]:
    flags: list[dict[str, str]] = []
    if not text.strip(): flags.append({"level": "review", "message": "No text extracted; manual review is required."})
    if not fields.get("name"): flags.append({"level": "review", "message": "Name was not detected; verify whether it is present in the source."})
    if kind == "invoice" and not fields.get("total"):
        flags.append({"level": "review", "message": "Invoice total was not detected; check the original document."})
    if kind == "resume" and not fields.get("qualifications"):
        flags.append({"level": "review", "message": "No listed qualification matched the starter patterns; review manually."})
    dates = fields.get("dates", [])
    if len(dates) > 1:
        parsed = []
        for date in dates:
            for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y", "%Y-%m-%d", "%d/%m/%y", "%m/%d/%y"):
                try: parsed.append((date, datetime.strptime(date, fmt).date())); break
                except ValueError: pass
        if len(parsed) >= 2 and max(d for _, d in parsed).year - min(d for _, d in parsed).year > 100:
            flags.append({"level": "review", "message": "Extracted dates span over 100 years; verify date formats and source values."})
    if re.search(r"\b(?:total|amount|राशि)\b", text, re.I) and len(fields.get("amounts", [])) > 1:
        flags.append({"level": "review", "message": "Multiple amounts detected; confirm which amount is the relevant total."})
    return flags
