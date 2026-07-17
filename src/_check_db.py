import chromadb
c = chromadb.PersistentClient(path=r"C:\Users\sinam\Desktop\ABC-RAG\ABC-RAG\data\chroma_db")
col = c.get_collection("yes24_books")
print("Count:", col.count())
sample = col.peek(2)
print("Sample IDs:", sample["ids"])
print("Sample metadatas:", sample["metadatas"][0] if sample["metadatas"] else "none")
