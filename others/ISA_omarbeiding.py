# pip install pdfplumber openai faiss-cpu tiktoken
import pdfplumber, re, openai, faiss, tiktoken, json, pathlib

PDF = "ISA 230 Revisjonsdokumentasjon.pdf"
MODEL = "text-embedding-3-small"          # billig & rask
INDEX_FILE = "isa_index.faiss"
META_FILE  = "isa_meta.jsonl"

tokenizer = tiktoken.get_encoding("cl100k_base")       # samme som OpenAI-mod.

def extract_chunks(pdf_path, max_tokens=250):
    chunks = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            raw = page.extract_text() or ""
            # Del på linjeskift, behold nummererte avsnitt samlet
            for para in re.split(r"\n{2,}", raw):
                para = para.strip()
                if not para:
                    continue
                # Hent seksjons-ID om mulig
                m = re.match(r"^([A]?\d+(\.\d+)*)(\s|-)\s*(.*)", para)
                sec = m.group(1) if m else None
                txt = para if len(tokenizer.encode(para)) <= max_tokens else \
                      " ".join(para.split()[:max_tokens])
                chunks.append({"text": txt, "section": sec})
    return chunks

def embed(text_list):
    resp = openai.Embedding.create(model=MODEL, input=text_list)
    return [d["embedding"] for d in resp["data"]]

def build_index(chunks):
    xb = faiss.IndexFlatL2(1536)  # 1536-dim for text-embedding-3-small
    batch, metas = [], []
    for ch in chunks:
        batch.append(ch["text"])
        metas.append(ch)
        # embed i batcher på 96 el.l.
        if len(batch) == 96:
            xb.add(np.array(embed(batch)).astype("float32"))
            batch = []
    if batch:
        xb.add(np.array(embed(batch)).astype("float32"))
    faiss.write_index(xb, INDEX_FILE)
    with open(META_FILE, "w", encoding="utf-8") as f:
        for m in metas:
            f.write(json.dumps(m, ensure_ascii=False) + "\n")

if __name__ == "__main__":
    chunks = extract_chunks(PDF)
    build_index(chunks)
    print(f"Indeksert {len(chunks)} biter fra {PDF}")
