# ```python
import os
import glob
import json
import uuid
import asyncio

from supabase import create_client, Client
from openai import AsyncOpenAI

# ── CONFIG ─────────────────────────────────────────────────────────────────────
SUPABASE_URL = os.getenv("SUPABASE_URL", "https://your-project.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "your-service-role-key")

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "sk-…")
EMBEDDING_MODEL = "text-embedding-ada-002"

INPUT_GLOB = "output/*.jsonl"

# ── CLIENTS ────────────────────────────────────────────────────────────────────
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
client = AsyncOpenAI(api_key=OPENAI_API_KEY)

# ── HELPERS ────────────────────────────────────────────────────────────────────
def load_jsonl(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.loads(f.readline())

async def get_embedding(text: str):
    response = await client.embeddings.create(
        model=EMBEDDING_MODEL,
        input=[text]
    )
    return response.data[0].embedding

# ── MAIN ───────────────────────────────────────────────────────────────────────
async def main():
    for filepath in glob.glob(INPUT_GLOB):
        record = load_jsonl(filepath)
        source = record.get("source")
        blocks = record.get("blocks", [])

        # Flatten blocks into one content string
        content_parts = []
        for b in blocks:
            if b.get("type") == "text":
                content_parts.append(b.get("content", ""))
            elif b.get("type") == "image_ocr":
                content_parts.append(f"[IMAGE {b.get('filename')} OCR]\n{b.get('content')}")
            elif b.get("type") == "table":
                rows = ["\t".join(row) for row in b.get("content", [])]
                content_parts.append("[TABLE]\n" + "\n".join(rows) + "\n[/TABLE]")
        content = "\n\n".join(content_parts)

        # Generate embedding
        embedding = await get_embedding(content)

        # Prepare metadata
        metadata = {"source": source, "blocks": blocks}

        # Insert into Supabase (returns full representation by default)
        try:
            response = (
                supabase
                    .table("n8n_test")
                    .insert({
                        "content": content,
                        "metadata": metadata,
                        "embedding": embedding
                    })
                    .execute()
            )
            new_id = response.data[0].get("id")
            print(f"Inserted {source} (id={new_id})")
        except Exception as e:
            print(f"Error inserting {source}")
            break

if __name__ == "__main__":
    asyncio.run(main())
# ```
