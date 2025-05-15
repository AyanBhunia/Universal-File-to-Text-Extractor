import os
import glob
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
from bs4 import BeautifulSoup     # for parsing HTML in ZIPs

# Setup paths
temps = {
    'images': 'temp_images'
}
INPUT_DIR  = 'input/'
OUTPUT_DIR = 'output/'

# Ensure directories exist
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(temps['images'], exist_ok=True)

# Utility: OCR image
def ocr_image(path):
    img = Image.open(path)
    try:
        txt = pytesseract.image_to_string(img).strip()
    finally:
        img.close()
    return txt

#
# DOCX extractor (unchanged)
#
def iter_block_items(parent):
    for child in parent.element.body:
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)

def extract_docx(path, img_dir):
    # clear images
    for f in os.listdir(img_dir):
        os.remove(os.path.join(img_dir, f))
    docx2txt.process(path, img_dir)
    image_files = sorted(glob.glob(os.path.join(img_dir, '*')))

    doc = Document(path)
    lines = []
    img_idx = 0
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            txt = ''
            for run in block.runs:
                if run._element.xpath('.//w:drawing'):
                    if txt.strip():
                        lines.append(txt); txt = ''
                    if img_idx < len(image_files):
                        img_path = image_files[img_idx]
                        lines.append(f"[IMAGE: {os.path.basename(img_path)}]")
                        lines.append(ocr_image(img_path))
                        img_idx += 1
                else:
                    txt += run.text
            if txt.strip():
                lines.append(txt)
            lines.append('')
        else:  # Table
            lines.append('[TABLE]')
            for row in block.rows:
                lines.append('\t'.join(c.text for c in row.cells))
            lines.append('[/TABLE]\n')
    return '\n'.join(lines)

#
# PDF extractor (PyMuPDF → inline text & images)
#
# PDF extractor (PyMuPDF → inline text & images)
def extract_pdf(path, img_dir):
    # wipe temp images
    for f in os.listdir(img_dir):
        os.remove(os.path.join(img_dir, f))

    doc = fitz.open(path)
    lines = []
    img_count = 0

    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if b["type"] == 0:  # text block
                for line in b["lines"]:
                    txt = "".join(span["text"] for span in line["spans"])
                    if txt.strip():
                        lines.append(txt)
                lines.append("")
            elif b["type"] == 1:  # image block
                img_count += 1

                # first try to extract by xref, else use embedded bytes
                if "xref" in b:
                    img_info = doc.extract_image(b["xref"])
                else:
                    img_info = {
                        "image": b["image"],
                        "ext": b.get("ext", "png")
                    }

                fname = f"pdf_img_{img_count}.{img_info['ext']}"
                path_img = os.path.join(img_dir, fname)
                with open(path_img, "wb") as imgf:
                    imgf.write(img_info["image"])

                # inline the OCR
                lines.append(f"[IMAGE: {fname}]")
                lines.append(ocr_image(path_img))
                lines.append("")

                # clean up immediately
                os.remove(path_img)
    doc.close()
    return "\n".join(lines)


#
# RTF extractor (text only; images in RTF are rare)
#
def extract_rtf(path, _img_dir=None):
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        raw = f.read()
    return rtf_to_text(raw)

#
# TXT extractor (passthrough)
#
def extract_txt(path, _img_dir=None):
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()

#
# ZIP (HTML) extractor: parse each HTML, interleave text and <img> OCR
#
def extract_zip_html(path, img_dir):
    # unzip to temp folder
    temp_dir = path + "_unzipped"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    with zipfile.ZipFile(path, 'r') as z:
        z.extractall(temp_dir)

    contents = []
    for root, _, files in os.walk(temp_dir):
        for fn in sorted(files):
            if fn.lower().endswith((".html", ".htm")):
                full = os.path.join(root, fn)
                contents.append(f"== {os.path.relpath(full, temp_dir)} ==")

                html = open(full, 'r', encoding='utf-8', errors='ignore').read()
                soup = BeautifulSoup(html, "html.parser")

                # iterate through direct children of <body>
                for node in soup.body.children:
                    if getattr(node, "name", None) in ("p", "div", "span", "h1", "h2", "h3"):
                        txt = node.get_text(strip=True)
                        if txt:
                            contents.append(txt)
                    elif getattr(node, "name", None) == "img":
                        src = node.get("src", "")
                        img_path = os.path.join(root, src)
                        if os.path.isfile(img_path):
                            # copy to temp_images, OCR, then delete
                            dst = os.path.join(img_dir, os.path.basename(img_path))
                            shutil.copy(img_path, dst)
                            contents.append(f"[IMAGE: {os.path.basename(dst)}]")
                            contents.append(ocr_image(dst))
                            contents.append("")
                            os.remove(dst)
                    # skip other tags/text nodes silently

                contents.append("")

    # cleanup
    shutil.rmtree(temp_dir)
    return "\n".join(contents)

# Dispatcher
handlers = {
    '.docx': extract_docx,
    '.pdf':  extract_pdf,
    '.rtf':  extract_rtf,
    '.txt':  extract_txt,
    '.zip':  extract_zip_html,
}

if __name__ == '__main__':
    img_dir = temps['images']

    # initial cleanup
    for f in os.listdir(img_dir):
        os.remove(os.path.join(img_dir, f))

    for fp in glob.glob(os.path.join(INPUT_DIR, '*')):
        name, ext = os.path.splitext(fp)
        handler = handlers.get(ext.lower())
        if not handler:
            print(f"Skipping {fp}: unsupported extension {ext}")
            continue

        print(f"Processing {os.path.basename(fp)}...")
        out_txt = handler(fp, img_dir)

        # write out
        out_path = os.path.join(OUTPUT_DIR, os.path.basename(name) + ext + '.txt')
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(out_txt)

        # cleanup temp images again
        for f in os.listdir(img_dir):
            os.remove(os.path.join(img_dir, f))

    print("Done.")
