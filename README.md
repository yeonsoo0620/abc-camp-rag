# Yes24 IT 모바일 베스트셀러 대시보드 & RAG 챗봇

예스24 IT 모바일 종합 베스트셀러 도서 데이터를 수집·분석하고, 시각화 대시보드와 AI 기반 도서 추천 챗봇을 구축한 프로젝트입니다.

## 프로젝트 구조

```
ABC-RAG/
├── src/
│   ├── app.py                  # Streamlit 대시보드 앱 (메인)
│   ├── build_embeddings.py     # KLUE-BERT 임베딩 → ChromaDB 저장
│   ├── make_excel_dashboard.py # Excel 대시보드 생성
│   └── _check_db.py           # ChromaDB 연결 확인 유틸리티
├── data/
│   ├── yes24_it_mobile_bestsellers.csv  # 수집된 베스트셀러 데이터
│   ├── chroma_db/              # ChromaDB 벡터 데이터베이스
│   ├── metadata (1).tsv        # 메타데이터
│   └── vectors (1).tsv         # 벡터 데이터
├── scrape_yes24.py             # BeautifulSoup 기반 스크래퍼
├── yes24_scraper.py            # Scrapling 기반 스크래퍼 (대안)
├── implementation_plan.md      # 구현 계획서
├── requirements.txt            # Python 의존성
└── .streamlit/config.toml      # Streamlit 테마 설정
```

## 주요 기능

### 1. 웹 스크래핑 (`scrape_yes24.py`)
- 예스24 IT 모바일 베스트셀러 전체 페이지 자동 수집
- 도서명, 저자, 출판사, 출판일, 판매가, 정가, 할인율, 판매지수 등 11개 필드 추출
- 서버 부하 방지를 위한 1초 요청 간격 적용
- UTF-8-SIG 인코딩으로 Excel 호환 CSV 저장

### 2. Streamlit 대시보드 (`src/app.py`)
8개 탭으로 구성된 인터랙티브 분석 대시보드:

| 탭 | 내용 |
|---|---|
| 개요 | 연도별 출판 수, 가격대별 분포, 데이터 미리보기 |
| 가격 분석 | 히스토그램, 출판사별 평균 판매가, 정가 vs 판매가 산점도 |
| 출판사 분석 | 도서 수, 평균 판매지수, 종합 통계 테이블 |
| 할인율 분석 | 할인율 분포, 출판사별 평균 할인율, 할인율 vs 판매지수 |
| 판매지수 분석 | Top 20 도서, 출판사별 총 판매지수, 판매가 vs 판매지수 |
| 트렌드 분석 | 연도별 판매가/판매지수 추이, 출판사 점유율 변화 |
| 키워드 검색 | 제목·저자·출판사 키워드 검색, 하이라이트, 인기 키워드 |
| 도서 추천 챗봇 | RAG 기반 AI 추천 (하단 설명) |

**사이드바 필터**: 판매가 범위, 출판사, 출판 연도, 최소 순위, 할인율 범위

### 3. RAG 도서 추천 챗봇
- **임베딩 모델**: KLUE-BERT (`klue/bert-base`)
- **벡터 DB**: ChromaDB (cosine 유사도)
- **LLM**: Groq API (Llama 3.3 70B / Llama 3.1 8B / Gemma 2 9B)
- **Function Calling**: 가격 필터, 판매지수 필터 함수 연동
- IT/모바일 도서 전문 추천, 외부 주제 요청 시 안내 메시지 반환

### 4. Excel 대시보드 (`src/make_excel_dashboard.py`)
- 5개 시트: 대시보드, 가격 분석, 출판사 분석, 할인율 분석, 원본 데이터
- KPI 카드, 차트, 조건부 서식 자동 생성

## 설치 및 실행

### 환경 설정

```bash
# 가상환경 생성
uv venv
uv pip install -r requirements.txt
```

### 의존성

| 패키지 | 버전 |
|---|---|
| streamlit | >= 1.30.0 |
| pandas | >= 2.0.0 |
| plotly | >= 5.18.0 |
| openpyxl | >= 3.1.0 |
| chromadb | >= 0.4.0 |
| sentence-transformers | >= 2.2.0 |

### 데이터 수집

```bash
# BeautifulSoup 기반 스크래퍼 실행
python scrape_yes24.py
```

### 임베딩 구축 (RAG 챗봇 필요 시)

```bash
# CSV → KLUE-BERT 임베딩 → ChromaDB 저장
python src/build_embeddings.py
```

### 대시보드 실행

```bash
# Streamlit 대시보드 실행
streamlit run src/app.py
```

> RAG 챗봇 탭 사용 시 사이드바에서 Groq API Key를 입력해야 합니다.
> 무료 발급: https://console.groq.com

## 데이터 수집 필드

| 필드 | 설명 |
|---|---|
| 순위 | 베스트셀러 순위 |
| 도서번호 | 예스24 상품 고유 번호 |
| 도서명 | 책 제목 |
| 링크 | 예스24 상세 페이지 URL |
| 저자 | 저자명 |
| 출판사 | 출판사명 |
| 출판일 | 출판 일자 |
| 판매가 | 현재 판매 가격 |
| 정가 | 원래 가격 |
| 할인율 | 할인 비율 |
| 판매지수 | 예스24 판매 지수 |
