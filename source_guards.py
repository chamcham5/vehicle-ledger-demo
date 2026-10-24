"""Text extraction guards for the bundled fictional PDF schema only."""
import re
from pypdf import PdfReader

def read_pdf_pages(path, marker):
    reader = PdfReader(path)
    if reader.is_encrypted:
        raise ValueError("Encrypted PDF unsupported by this probe")
    pages = []
    for number, page in enumerate(reader.pages, 1):
        text = page.extract_text()
        if not text or marker not in text:
            raise ValueError(f"Missing text/schema in {path.name} page{number}")
        pages.append((number, text))
    if not pages:
        raise ValueError("No readable PDF pages")
    return pages

def field(text, label):
    matches = re.findall(r"^" + re.escape(label) + r": (.+)$", text, re.M)
    if len(matches) != 1:
        raise ValueError(f"Expected exactly one {label}")
    return matches[0].strip()

def jint(value):
    if not re.fullmatch(r"0|[1-9]\d*", str(value)):
        raise ValueError("Invalid nonnegative integer JPY")
    return int(value)
