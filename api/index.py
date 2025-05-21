from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from starlette.responses import JSONResponse
from pathlib import Path
from extractors.handlers import DISPATCH
import uuid
import os
import json
from typing import List

app = FastAPI()

@app.post("/extract")
async def extract(
    mode: str = Form(..., pattern="^(single|multiple)$"),
    output_type: str = Form(..., pattern="^(text|jsonl|blocks)$"),
    files: List[UploadFile] = File(...)
):
    if mode == "single" and len(files) != 1:
        raise HTTPException(400, "single mode requires exactly one file")
    if mode == "multiple" and len(files) < 1:
        raise HTTPException(400, "multiple mode requires at least one file")

    results = []
    for file in files:
        suffix = Path(file.filename).suffix.lower()
        handler = DISPATCH.get(suffix)
        if not handler:
            continue  # skip unsupported files
        tmp = f"/tmp/{uuid.uuid4()}{suffix}"
        with open(tmp, "wb") as f:
            f.write(await file.read())
        blocks = handler(tmp)
        results.append({
            "id": str(uuid.uuid4()),
            "source": file.filename,
            "blocks": blocks,
        })
        try:
            os.remove(tmp)
        except Exception:
            pass

    # Output in requested format
    if output_type == "jsonl":
        # Each file becomes a single JSONL line string
        data = [json.dumps(r, ensure_ascii=False) for r in results]
        return JSONResponse(content={"data": data})

    elif output_type == "text":
        # Concatenate all text (flatten) from all blocks
        all_texts = []
        for r in results:
            texts = []
            for blk in r["blocks"]:
                if blk["type"] == "text":
                    texts.append(blk["content"])
                elif blk["type"].endswith("image_ocr"):
                    texts.append(f"[{blk.get('filename','IMAGE')}] {blk['content']}")
                elif blk["type"] == "table":
                    texts.append("\n".join(["\t".join(row) for row in blk["content"]]))
                elif blk["type"] == "meta":
                    texts.append(blk["content"])
            all_texts.append({
                "id": r["id"], 
                "source": r["source"], 
                "text": "\n\n".join(texts)
            })
        return JSONResponse(content={"data": all_texts})

    # Default, "blocks" format: full JSON array of extracted (id, source, blocks) per file
    return JSONResponse(content={"data": results})

@app.get("/test")
def test_connection():
    return {"message": "Connected"}

@app.get("/test-ocr")
async def test_ocr():
    """Test if the OCR binary works"""
    from PIL import Image, ImageDraw, ImageFont
    import pytesseract
    import os
    from pathlib import Path
    
    # Create a test image with text
    img = Image.new('RGB', (200, 50), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    d.text((10,10), "Testing OCR", fill=(0,0,0))
    test_img_path = "/tmp/test_ocr.png"
    img.save(test_img_path)
    
    # Get paths
    binary_path = os.path.join(os.getcwd(), "api", "ocr-bin", "bin", "libtesseract64.so")
    tessdata_path = os.path.join(os.getcwd(), "api", "ocr-bin", "share", "libtesseract64.so")
    
    return {
        "binary_exists": os.path.exists(binary_path),
        "tessdata_exists": os.path.exists(tessdata_path),
        "language_files": os.listdir(tessdata_path) if os.path.exists(tessdata_path) else [],
        "ocr_result": pytesseract.image_to_string(Image.open(test_img_path)).strip()
    }