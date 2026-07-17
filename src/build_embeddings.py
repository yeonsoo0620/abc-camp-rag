"""CSV 데이터를 KLUE-BERT 임베딩하여 ChromaDB에 저장하는 스크립트."""

import pandas as pd
import chromadb
from sentence_transformers import SentenceTransformer
import os

DATA_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "yes24_it_mobile_bestsellers.csv")
DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "chroma_db")
COLLECTION_NAME = "yes24_books"
MODEL_NAME = "klue/bert-base"


def build_document(row):
    return (
        f"{row['도서명']} | 저자: {row['저자']} | 출판사: {row['출판사']} | "
        f"출판일: {row['출판일']} | 판매가: {row['판매가']}원 | "
        f"정가: {row['정가']}원 | 할인율: {row['할인율']} | "
        f"판매지수: {row['판매지수']} | 링크: {row['링크']}"
    )


def main():
    print("[1/4] CSV 데이터 로드 중...")
    df = pd.read_csv(DATA_PATH, encoding="utf-8-sig")
    print(f"  -> {len(df)}권 로드 완료")

    print("[2/4] KLUE-BERT 모델 로드 중...")
    model = SentenceTransformer(MODEL_NAME)
    print("  -> 모델 로드 완료")

    print("[3/4] ChromaDB 클라이언트 초기화 중...")
    client = chromadb.PersistentClient(path=DB_PATH)

    try:
        client.delete_collection(COLLECTION_NAME)
        print("  -> 기존 컬렉션 삭제 완료")
    except Exception:
        pass

    collection = client.create_collection(
        name=COLLECTION_NAME,
        metadata={"hnsw:space": "cosine"},
    )

    print("[4/4] 임베딩 생성 및 저장 중...")
    documents = []
    metadatas = []
    ids = []

    for _, row in df.iterrows():
        doc = build_document(row)
        documents.append(doc)
        metadatas.append({
            "순위": int(row["순위"]) if pd.notna(row["순위"]) else 0,
            "도서명": str(row["도서명"]),
            "저자": str(row["저자"]),
            "출판사": str(row["출판사"]),
            "출판일": str(row["출판일"]),
            "판매가": str(row["판매가"]),
            "정가": str(row["정가"]),
            "할인율": str(row["할인율"]),
            "판매지수": str(row["판매지수"]),
            "링크": str(row["링크"]),
        })
        ids.append(f"book_{int(row['순위']) if pd.notna(row['순위']) else len(ids)}")

    batch_size = 100
    for i in range(0, len(documents), batch_size):
        batch_docs = documents[i:i + batch_size]
        batch_meta = metadatas[i:i + batch_size]
        batch_ids = ids[i:i + batch_size]

        embeddings = model.encode(batch_docs, show_progress_bar=False).tolist()

        collection.add(
            documents=batch_docs,
            embeddings=embeddings,
            metadatas=batch_meta,
            ids=batch_ids,
        )
        print(f"  -> {min(i + batch_size, len(documents))}/{len(documents)} 저장 완료")

    print(f"\n[DONE] ChromaDB 저장 완료: {collection.count()}권")
    print(f"  DB 경로: {DB_PATH}")


if __name__ == "__main__":
    main()
