import os
import glob
import uuid
import shutil
import zipfile

import docx2txt
from pathlib import Path
import pytesseract
from PIL import Image
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

ROOT = Path(__file__).resolve().parent.parent / "api" / "ocr-bin"
os.environ["PATH"] = str(ROOT / "bin") + os.pathsep + os.environ["PATH"]
os.environ["TESSDATA_PREFIX"] = str(ROOT / "share/tessdata")
pytesseract.pytesseract.tesseract_cmd = str(ROOT / "bin" / "tesseract.exe")

import fitz                       # PyMuPDF
from striprtf.striprtf import rtf_to_text
from bs4 import BeautifulSoup     # HTML parsing

IMG_DIR = "/tmp/images"
os.makedirs(IMG_DIR, exist_ok=True)

# ——— Utility: OCR image path ———
def ocr_image_path(img_path):
    img = Image.open(img_path)
    try:
        return pytesseract.image_to_string(img).strip()
    finally:
        img.close()

# ——— DOCX: block extractor ———
def iter_block_items(parent):
    for child in parent.element.body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            yield Paragraph(child, parent)
        elif tag == 'tbl':
            yield Table(child, parent)

def extract_docx(path: str) -> list:
    # clear temp images
    for f in os.listdir(IMG_DIR):
        os.remove(os.path.join(IMG_DIR, f))
    docx2txt.process(path, IMG_DIR)
    image_files = sorted(glob.glob(os.path.join(IMG_DIR, "*")))

    doc = Document(path)
    blocks = []
    img_ix = 0
    for blk in iter_block_items(doc):
        if isinstance(blk, Paragraph):
            accum = ""
            for run in blk.runs:
                if run._element.xpath('.//w:drawing'):
                    if accum.strip():
                        blocks.append({"type": "text", "content": accum.strip()})
                        accum = ""
                    if img_ix < len(image_files):
                        img_path = image_files[img_ix]
                        blocks.append({
                            "type": "image_ocr",
                            "filename": os.path.basename(img_path),
                            "content":  ocr_image_path(img_path)
                        })
                        img_ix += 1
                else:
                    accum += run.text
            if accum.strip():
                blocks.append({"type": "text", "content": accum.strip()})
        else:  # Table
            rows = [[cell.text for cell in row.cells] for row in blk.rows]
            blocks.append({"type": "table", "content": rows})
    # Clean temp images? Optional.
    for f in os.listdir(IMG_DIR):
        os.remove(os.path.join(IMG_DIR, f))
    return blocks

# ——— PDF: block extractor ———
def extract_pdf(path: str) -> list:
    for f in os.listdir(IMG_DIR):
        os.remove(os.path.join(IMG_DIR, f))
    doc = fitz.open(path)
    blocks = []
    img_ct = 0
    for page in doc:
        for b in page.get_text("dict")["blocks"]:
            if b["type"] == 0:  # text
                for line in b["lines"]:
                    txt = "".join(span["text"] for span in line["spans"]).strip()
                    if txt:
                        blocks.append({"type": "text", "content": txt})
            elif b["type"] == 1:  # image
                img_ct += 1
                if "xref" in b:
                    info = doc.extract_image(b["xref"])
                else:
                    info = {"image": b["image"], "ext": b.get("ext", "png")}
                name = f"pdf_img_{img_ct}.{info['ext']}"
                dst  = os.path.join(IMG_DIR, name)
                with open(dst, "wb") as out:
                    out.write(info["image"])
                blocks.append({
                    "type":     "image_ocr",
                    "filename": name,
                    "content":  ocr_image_path(dst)
                })
                os.remove(dst)
    doc.close()
    return blocks

# ——— RTF: block extractor ———
def extract_rtf(path: str) -> list:
    txt = open(path, "r", encoding="utf-8", errors="ignore").read()
    lines = [l.strip() for l in rtf_to_text(txt).splitlines() if l.strip()]
    return [{"type": "text", "content": l} for l in lines]

# ——— TXT: block extractor ———
def extract_txt(path: str) -> list:
    lines = [l.strip() for l in open(path, "r", encoding="utf-8", errors="ignore").read().splitlines() if l.strip()]
    return [{"type": "text", "content": l} for l in lines]

# ——— ZIP (HTMLs): block extractor ———
def extract_zip(path: str) -> list:
    temp_dir = path + "_unzipped"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    with zipfile.ZipFile(path, "r") as z:
        z.extractall(temp_dir)
    blocks = []
    for root, _, files in os.walk(temp_dir):
        for fn in sorted(files):
            if fn.lower().endswith((".html", ".htm")):
                full = os.path.join(root, fn)
                blocks.append({"type": "meta", "content": f"== {os.path.relpath(full, temp_dir)} =="})
                html = open(full, "r", encoding="utf-8", errors="ignore").read()
                soup = BeautifulSoup(html, "html.parser")
                for node in (soup.body.children if soup.body else []):
                    if getattr(node, "name", None) in ("p","div","span","h1","h2","h3"):
                        txt = node.get_text(strip=True)
                        if txt:
                            blocks.append({"type": "text", "content": txt})
                    elif getattr(node, "name", None) == "img":
                        src = node.get("src","")
                        fpath = os.path.join(root, src)
                        if os.path.isfile(fpath):
                            name = os.path.basename(fpath)
                            dst  = os.path.join(IMG_DIR, name)
                            shutil.copy(fpath, dst)
                            blocks.append({
                                "type":     "image_ocr",
                                "filename": name,
                                "content":  ocr_image_path(dst)
                            })
                            os.remove(dst)
    shutil.rmtree(temp_dir)  # always clean up
    return blocks

# ——— Dispatcher ———
DISPATCH = {
    ".docx": extract_docx,
    ".pdf":  extract_pdf,
    ".rtf":  extract_rtf,
    ".txt":  extract_txt,
    ".zip":  extract_zip,
}

# Optional: Standalone CLI
if __name__ == "__main__":
    import json
    INPUT_DIR = "input/"
    OUTPUT_DIR = "output/"
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for filepath in glob.glob(os.path.join(INPUT_DIR, "*")):
        base, ext = os.path.splitext(os.path.basename(filepath))
        func = DISPATCH.get(ext.lower())
        if not func:
            print(f"Skipping {filepath} (unsupported: {ext})")
            continue
        print(f"Processing {base}{ext} → JSONL")
        blocks = func(filepath)
        record = {
            "id":     str(uuid.uuid4()),
            "source": base + ext,
            "blocks": blocks
        }
        outpath = os.path.join(OUTPUT_DIR, f"{base}{ext}.jsonl")
        with open(outpath, "w", encoding="utf-8") as out:
            out.write(json.dumps(record, ensure_ascii=False) + "\n")
    print("All done — JSONL records in", OUTPUT_DIR)