
import os
import glob
import docx2txt
import pytesseract
from PIL import Image
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn

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
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)


def extract_images(docx_path, image_dir):
    """Extract images using docx2txt and return sorted image filenames."""
    # Clear previous images
    for f in os.listdir(image_dir):
        os.remove(os.path.join(image_dir, f))
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
    # Extract all images first
    image_files = extract_images(docx_path, image_dir)
    img_index = 0

    doc = Document(docx_path)
    output_lines = []

    # Iterate through paragraphs and tables in sequence
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            # Process runs in paragraph
            line = ''
            for run in block.runs:
                # Check for image in run
                if run._element.xpath('.//w:drawing'):
                    # Flush any accumulated text
                    if line.strip():
                        output_lines.append(line)
                        line = ''
                    # OCR this image
                    if img_index < len(image_files):
                        img_path = image_files[img_index]
                        ocr_txt = ocr_image(img_path)
                        output_lines.append(f"[IMAGE OCR: {os.path.basename(img_path)}]")
                        output_lines.append(ocr_txt)
                        img_index += 1
                else:
                    line += run.text
            if line.strip():
                output_lines.append(line)
            output_lines.append('')  # blank line after paragraph

        elif isinstance(block, Table):
            # Process a table
            output_lines.append('[TABLE]')
            for row in block.rows:
                row_text = '\t'.join(cell.text for cell in row.cells)
                output_lines.append(row_text)
            output_lines.append('[/TABLE]\n')

    return '\n'.join(output_lines)

# Main processing loop
for filepath in glob.glob(os.path.join(INPUT_DIR, '*.docx')):
    filename = os.path.basename(filepath)
    out_path = os.path.join(OUTPUT_DIR, os.path.splitext(filename)[0] + '.txt')
    print(f"Processing {filename}...")

    final_text = extract_docx_inline(filepath, IMG_DIR)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(final_text)

    # Cleanup image folder for next document
    for img in os.listdir(IMG_DIR):
        os.remove(os.path.join(IMG_DIR, img))
        
    break
print("Done. Outputs in 'output' folder.")

