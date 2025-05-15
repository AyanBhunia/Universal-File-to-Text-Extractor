import os
import glob
import uuid
import json
import shutil
import zipfile

import docx2txt
import pytesseract
from PIL import Image
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

import fitz                       # PyMuPDF
from striprtf.striprtf import rtf_to_text
from bs4 import BeautifulSoup     # HTML parsing

# ——— Paths ———
INPUT_DIR   = "input/"
OUTPUT_DIR  = "output/"
IMG_DIR     = "temp_images/"

for d in (INPUT_DIR, OUTPUT_DIR, IMG_DIR):
    os.makedirs(d, exist_ok=True)

# ——— OCR Utility ———
def ocr_image(path):
    img = Image.open(path)
    try:
        return pytesseract.image_to_string(img).strip()
    finally:
        img.close()

# ——— DOCX → blocks ———
def iter_block_items(parent):
    for child in parent.element.body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            yield Paragraph(child, parent)
        elif tag == 'tbl':
            yield Table(child, parent)

def extract_docx_blocks(path):
    # clear image folder
    for f in os.listdir(IMG_DIR):
        os.remove(os.path.join(IMG_DIR, f))
    # extract raw images
    docx2txt.process(path, IMG_DIR)
    image_files = sorted(glob.glob(os.path.join(IMG_DIR, "*")))

    doc    = Document(path)
    blocks = []
    img_ix = 0

    for blk in iter_block_items(doc):
        if isinstance(blk, Paragraph):
            accum = ""
            for run in blk.runs:
                if run._element.xpath('.//w:drawing'):
                    # flush text
                    if accum.strip():
                        blocks.append({"type": "text", "content": accum.strip()})
                        accum = ""
                    # flush image
                    if img_ix < len(image_files):
                        img_path = image_files[img_ix]
                        blocks.append({
                            "type":     "image_ocr",
                            "filename": os.path.basename(img_path),
                            "content":  ocr_image(img_path)
                        })
                        img_ix += 1
                else:
                    accum += run.text
            if accum.strip():
                blocks.append({"type": "text", "content": accum.strip()})

        else:  # Table
            rows = [[cell.text for cell in row.cells] for row in blk.rows]
            blocks.append({"type": "table", "content": rows})

    return blocks

# ——— PDF → blocks ———
def extract_pdf_blocks(path):
    # clear image folder
    for f in os.listdir(IMG_DIR):
        os.remove(os.path.join(IMG_DIR, f))

    doc    = fitz.open(path)
    blocks = []
    img_ct = 0

    for page in doc:
        for b in page.get_text("dict")["blocks"]:
            if b["type"] == 0:  # text block
                for line in b["lines"]:
                    txt = "".join(span["text"] for span in line["spans"]).strip()
                    if txt:
                        blocks.append({"type": "text", "content": txt})
            elif b["type"] == 1:  # image block
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
                    "content":  ocr_image(dst)
                })
                os.remove(dst)

    doc.close()
    return blocks

# ——— RTF → blocks (text only) ———
def extract_rtf_blocks(path):
    txt = open(path, "r", encoding="utf-8", errors="ignore").read()
    content = rtf_to_text(txt).splitlines()
    return [{"type": "text", "content": line} for line in content if line.strip()]

# ——— TXT → blocks (text only) ———
def extract_txt_blocks(path):
    lines = open(path, "r", encoding="utf-8", errors="ignore").read().splitlines()
    return [{"type": "text", "content": line} for line in lines if line.strip()]

# ——— ZIP (HTML) → blocks ———
def extract_zip_html_blocks(path):
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

                for node in soup.body.children:
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
                                "content":  ocr_image(dst)
                            })
                            os.remove(dst)

    shutil.rmtree(temp_dir)
    return blocks

# ——— Dispatcher ———
DISPATCH = {
    ".docx": extract_docx_blocks,
    ".pdf":  extract_pdf_blocks,
    ".rtf":  extract_rtf_blocks,
    ".txt":  extract_txt_blocks,
    ".zip":  extract_zip_html_blocks,
}

if __name__ == "__main__":
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
