import pandas as pd
import numpy as np
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

df = pd.read_csv(r'C:\Users\sinam\Desktop\ABC-RAG\ABC-RAG\data\yes24_it_mobile_bestsellers.csv', encoding='utf-8-sig')

# Clean numeric columns
for col in ['판매가', '정가', '판매지수']:
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '', regex=False).str.strip(), errors='coerce')

df['할인율_num'] = pd.to_numeric(df['할인율'].str.replace('%', '', regex=False).str.strip(), errors='coerce')
df['출판일_연도'] = pd.to_numeric(df['출판일'].str.extract(r'(\d{4})', expand=False), errors='coerce')
df['출판일_월'] = pd.to_numeric(df['출판일'].str.extract(r'(\d{2})월', expand=False), errors='coerce')

print('=== BASIC INFO ===')
print(f'Shape: {df.shape}')
print(f'Total books: {len(df)}')
print(f'Avg price: {df["판매가"].mean():,.0f}')
print(f'Median price: {df["판매가"].median():,.0f}')
print(f'Avg sales index: {df["판매지수"].mean():,.0f}')
print(f'Max sales index: {df["판매지수"].max():,.0f}')
print(f'Avg discount: {df["할인율_num"].mean():.1f}%')
print(f'Unique publishers: {df["출판사"].nunique()}')
print(f'Year range: {int(df["출판일_연도"].min())} - {int(df["출판일_연도"].max())}')
print()

print('=== TOP 10 PUBLISHERS ===')
pub = df['출판사'].value_counts().head(10)
for p, c in pub.items():
    print(f'  {p}: {c}')
print()

print('=== PRICE DISTRIBUTION ===')
bins = [0, 15000, 20000, 25000, 30000, 50000, 200000]
labels = ['under_15k', '15k_20k', '20k_25k', '25k_30k', '30k_50k', 'over_50k']
df['가격대'] = pd.cut(df['판매가'], bins=bins, labels=labels)
for l in labels:
    cnt = (df['가격대'] == l).sum()
    pct = cnt / len(df) * 100
    print(f'  {l}: {cnt} ({pct:.1f}%)')
print()

print('=== TOP 15 BY SALES INDEX ===')
top15 = df.nlargest(15, '판매지수')[['순위', '도서명', '저자', '출판사', '판매가', '판매지수', '출판일_연도']]
for _, r in top15.iterrows():
    try:
        rank_val = int(r['순위']) if pd.notna(r['순위']) else '?'
        title_val = str(r['도서명'])[:40]
        pub_val = str(r['출판사'])
        price_val = int(r['판매가']) if pd.notna(r['판매가']) else 0
        sales_val = int(r['판매지수']) if pd.notna(r['판매지수']) else 0
        year_val = int(r['출판일_연도']) if pd.notna(r['출판일_연도']) else '?'
        print(f'  #{rank_val} {title_val} | {pub_val} | {price_val:,} | idx:{sales_val:,} | {year_val}')
    except:
        pass
print()

print('=== YEAR DISTRIBUTION ===')
yr = df['출판일_연도'].dropna().astype(int).value_counts().sort_index()
for y, c in yr.items():
    print(f'  {y}: {c}')
print()

print('=== DISCOUNT DISTRIBUTION ===')
disc = df['할인율'].dropna().value_counts().sort_index()
for d, c in disc.items():
    print(f'  {d}: {c} ({c/len(df)*100:.1f}%)')
print()

print('=== TOP PUBLISHERS BY SALES INDEX ===')
pub_sales = df.groupby('출판사').agg(
    cnt=('도서명', 'count'),
    avg_idx=('판매지수', 'mean'),
    total_idx=('판매지수', 'sum'),
    avg_price=('판매가', 'mean')
).sort_values('total_idx', ascending=False).head(10)
for p, r in pub_sales.iterrows():
    print(f'  {p}: cnt={int(r["cnt"])}, avg_idx={int(r["avg_idx"]):,}, total={int(r["total_idx"]):,}, avg_price={int(r["avg_price"]):,}')
print()

print('=== CORRELATIONS ===')
print(f'  Price vs Sales Index: {df[["판매가","판매지수"]].dropna().corr().iloc[0,1]:.4f}')
print(f'  Discount vs Sales Index: {df[["할인율_num","판매지수"]].dropna().corr().iloc[0,1]:.4f}')
print(f'  Price vs Discount: {df[["판매가","할인율_num"]].dropna().corr().iloc[0,1]:.4f}')
print()

print('=== AI RELATED BOOKS ===')
ai_kw = ['AI', 'ChatGPT', 'GPT', 'claude', 'gemini', '바이브', 'LLM', '딥러닝', '머신러닝', '인공지능']
mask = df['도서명'].str.contains('|'.join(ai_kw), case=False, na=False)
ai_books = df[mask]
print(f'  AI books count: {len(ai_books)} ({len(ai_books)/len(df)*100:.1f}%)')
print(f'  AI avg sales index: {ai_books["판매지수"].mean():,.0f}')
print(f'  Non-AI avg sales index: {df[~mask]["판매지수"].mean():,.0f}')
print(f'  AI avg price: {ai_books["판매가"].mean():,.0f}')
print(f'  Non-AI avg price: {df[~mask]["판매가"].mean():,.0f}')
print()

print('=== TOP AUTHORS ===')
auth = df['저자'].value_counts().head(10)
for a, c in auth.items():
    print(f'  {a}: {c}')
print()

print('=== YEARLY AVG PRICE ===')
yp = df.groupby('출판일_연도')['판매가'].mean().dropna()
for y, p in yp.items():
    print(f'  {int(y)}: {int(p):,}')
print()

print('=== YEARLY AVG SALES INDEX ===')
ys = df.groupby('출판일_연도')['판매지수'].mean().dropna()
for y, s in ys.items():
    print(f'  {int(y)}: {int(s):,}')
print()

print('=== PRICE RANGE STATS ===')
for label in labels:
    sub = df[df['가격대'] == label]
    print(f'  {label}: avg_sales={int(sub["판매지数"].mean()) if len(sub) > 0 else 0:,}, count={len(sub)}')

print()
print('=== PUBLISHER MARKET SHARE (TOP 5) ===')
top5_pub = df['출판사'].value_counts().head(5)
others = len(df) - top5_pub.sum()
for p, c in top5_pub.items():
    print(f'  {p}: {c} ({c/len(df)*100:.1f}%)')
print(f'  Others: {others} ({others/len(df)*100:.1f}%)')
