import os
import glob
import uuid
import json
import docx2txt
import pytesseract
from PIL import Image
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

# Setup paths
INPUT_DIR = 'input/'
OUTPUT_DIR = 'output/'
IMG_DIR = 'temp_images/'

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

# Optional: set tesseract path explicitly
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to yield paragraphs and tables in document order
def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    """
    for child in parent.element.body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            yield Paragraph(child, parent)
        elif tag == 'tbl':
            yield Table(child, parent)


def extract_images(docx_path, image_dir):
    """Extract images using docx2txt and return sorted image filenames."""
    # Clear previous images
    for fname in os.listdir(image_dir):
        os.remove(os.path.join(image_dir, fname))
    docx2txt.process(docx_path, image_dir)
    return sorted(glob.glob(os.path.join(image_dir, '*')))


def ocr_image(img_path):
    """Perform OCR on a single image and return text."""
    try:
        return pytesseract.image_to_string(Image.open(img_path)).strip()
    except Exception as e:
        return f"[OCR FAILED: {e}]"


def extract_docx_inline(docx_path, image_dir):
    """Extract text, tables, and inline images (with OCR) preserving document order."""
    image_files = extract_images(docx_path, image_dir)
    img_index = 0

    doc = Document(docx_path)
    blocks = []

    # Iterate through paragraphs and tables in sequence
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = ''
            for run in block.runs:
                if run._element.xpath('.//w:drawing'):
                    if text.strip():
                        blocks.append({'type': 'text', 'content': text.strip()})
                        text = ''
                    if img_index < len(image_files):
                        img_path = image_files[img_index]
                        ocr_txt = ocr_image(img_path)
                        blocks.append({'type': 'image_ocr',
                                       'filename': os.path.basename(img_path),
                                       'content': ocr_txt})
                        img_index += 1
                else:
                    text += run.text
            if text.strip():
                blocks.append({'type': 'text', 'content': text.strip()})
        elif isinstance(block, Table):
            # Flatten table rows
            rows = []
            for row in block.rows:
                rows.append([cell.text for cell in row.cells])
            blocks.append({'type': 'table', 'content': rows})

    return blocks

# Main processing loop: output one JSONL line per document
for filepath in glob.glob(os.path.join(INPUT_DIR, '*.docx')):
    filename = os.path.basename(filepath)
    base, _ = os.path.splitext(filename)
    out_path = os.path.join(OUTPUT_DIR, f'{base}.jsonl')
    print(f"Processing {filename}...")

    blocks = extract_docx_inline(filepath, IMG_DIR)
    record = {
        'id': str(uuid.uuid4()),
        'source': filename,
        'blocks': blocks
    }

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")
    
    # Cleanup image folder for next document
    for img in os.listdir(IMG_DIR):
        os.remove(os.path.join(IMG_DIR, img))
    break

print("Done. JSONL files in 'output' folder.")
