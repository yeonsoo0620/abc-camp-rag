import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Try to find Korean font
font_path = None
for fp in fm.findSystemFonts():
    if 'malgun' in fp.lower() or 'gothic' in fp.lower() or 'nanum' in fp.lower():
        font_path = fp
        break

if font_path:
    font_prop = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
    print(f'Font: {font_path}')
else:
    plt.rcParams['font.family'] = 'DejaVu Sans'
    print('Using DejaVu Sans')

plt.rcParams['axes.unicode_minus'] = False

df = pd.read_csv(r'C:\Users\sinam\Desktop\ABC-RAG\ABC-RAG\data\yes24_it_mobile_bestsellers.csv', encoding='utf-8-sig')

for col in ['판매가', '정가', '판매지수']:
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '', regex=False).str.strip(), errors='coerce')
df['할인율_num'] = pd.to_numeric(df['할인율'].str.replace('%', '', regex=False).str.strip(), errors='coerce')
df['출판일_연도'] = pd.to_numeric(df['출판일'].str.extract(r'(\d{4})', expand=False), errors='coerce')

OUT = r'C:\Users\sinam\Desktop\ABC-RAG\ABC-RAG\charts'
os.makedirs(OUT, exist_ok=True)

# Color palette
COLORS = ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe', '#00f2fe', '#43e97b', '#38f9d7', '#fa709a', '#fee140']
BG = '#fafbfc'

# 1. Top 15 Sales Index
fig, ax = plt.subplots(figsize=(12, 7), facecolor=BG)
ax.set_facecolor(BG)
top15 = df.nlargest(15, '판매지수').sort_values('판매지수')
bars = ax.barh(range(15), top15['판매지수'], color=COLORS[:15], edgecolor='white', height=0.7)
ax.set_yticks(range(15))
labels = [f"#{int(r['순위'])} {str(r['도서명'])[:25]}..." if len(str(r['도서명']))>25 else f"#{int(r['순위'])} {r['도서명']}" for _, r in top15.iterrows()]
ax.set_yticklabels(labels, fontsize=9)
for i, (bar, val) in enumerate(zip(bars, top15['판매지수'])):
    ax.text(bar.get_width() + 500, bar.get_y() + bar.get_height()/2, f'{int(val):,}', va='center', fontsize=9, fontweight='bold')
ax.set_xlabel('Sales Index', fontsize=11)
ax.set_title('Top 15 Books by Sales Index', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/01_top15_sales_index.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 1 done')

# 2. Price Distribution
fig, ax = plt.subplots(figsize=(10, 6), facecolor=BG)
ax.set_facecolor(BG)
bins = [0, 15000, 20000, 25000, 30000, 50000, 200000]
labels_p = ['<15K', '15-20K', '20-25K', '25-30K', '30-50K', '>50K']
df['price_bin'] = pd.cut(df['판매가'], bins=bins, labels=labels_p)
counts = df['price_bin'].value_counts().reindex(labels_p)
bars = ax.bar(labels_p, counts.values, color=COLORS[:6], edgecolor='white', width=0.65)
for bar, val in zip(bars, counts.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5, f'{val}\n({val/len(df)*100:.1f}%)', ha='center', fontsize=9, fontweight='bold')
ax.set_xlabel('Price Range (KRW)', fontsize=11)
ax.set_ylabel('Number of Books', fontsize=11)
ax.set_title('Price Distribution of Bestsellers', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/02_price_distribution.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 2 done')

# 3. Publisher Market Share (Pie)
fig, ax = plt.subplots(figsize=(10, 8), facecolor=BG)
top5 = df['출판사'].value_counts().head(5)
others_count = len(df) - top5.sum()
pie_data = pd.concat([top5, pd.Series({'Others': others_count})])
colors_pie = COLORS[:6]
wedges, texts, autotexts = ax.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', colors=colors_pie,
                                   startangle=90, pctdistance=0.8, textprops={'fontsize': 10})
for at in autotexts:
    at.set_fontweight('bold')
ax.set_title('Publisher Market Share (Top 5)', fontsize=14, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig(f'{OUT}/03_publisher_share.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 3 done')

# 4. Year Distribution
fig, ax = plt.subplots(figsize=(12, 6), facecolor=BG)
ax.set_facecolor(BG)
yr = df['출판일_연도'].dropna().astype(int).value_counts().sort_index()
recent = yr[yr.index >= 2020]
bars = ax.bar(recent.index.astype(str), recent.values, color='#667eea', edgecolor='white', width=0.6)
for bar, val in zip(bars, recent.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 3, str(val), ha='center', fontsize=10, fontweight='bold')
ax.set_xlabel('Year', fontsize=11)
ax.set_ylabel('Number of Books', fontsize=11)
ax.set_title('Publication Year Distribution (2020-2026)', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/04_year_distribution.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 4 done')

# 5. AI vs Non-AI Sales Index
fig, ax = plt.subplots(figsize=(10, 6), facecolor=BG)
ax.set_facecolor(BG)
ai_kw = ['AI', 'ChatGPT', 'GPT', 'claude', 'gemini', '바이브', 'LLM', '딥러닝', '머신러닝', '인공지능']
mask = df['도서명'].str.contains('|'.join(ai_kw), case=False, na=False)
ai_avg = df[mask]['판매지수'].mean()
non_ai_avg = df[~mask]['판매지수'].mean()
bars = ax.bar(['AI/LLM Related', 'Non-AI/LLM'], [ai_avg, non_ai_avg], color=['#f5576c', '#667eea'], edgecolor='white', width=0.5)
for bar, val in zip(bars, [ai_avg, non_ai_avg]):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50, f'{val:,.0f}', ha='center', fontsize=12, fontweight='bold')
ax.set_ylabel('Avg Sales Index', fontsize=11)
ax.set_title('AI/LLM Books vs Non-AI Books: Avg Sales Index', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/05_ai_vs_nonai.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 5 done')

# 6. Top Publishers by Total Sales Index
fig, ax = plt.subplots(figsize=(12, 7), facecolor=BG)
ax.set_facecolor(BG)
pub_total = df.groupby('출판사')['판매지수'].sum().nlargest(10).sort_values()
bars = ax.barh(range(10), pub_total.values, color=COLORS[:10], edgecolor='white', height=0.65)
ax.set_yticks(range(10))
ax.set_yticklabels(pub_total.index, fontsize=10)
for bar, val in zip(bars, pub_total.values):
    ax.text(bar.get_width() + 2000, bar.get_y() + bar.get_height()/2, f'{int(val):,}', va='center', fontsize=9, fontweight='bold')
ax.set_xlabel('Total Sales Index', fontsize=11)
ax.set_title('Top 10 Publishers by Total Sales Index', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/06_publisher_sales_total.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 6 done')

# 7. Price vs Sales Index Scatter
fig, ax = plt.subplots(figsize=(10, 7), facecolor=BG)
ax.set_facecolor(BG)
scatter_data = df[['판매가', '판매지수', '도서명']].dropna()
ax.scatter(scatter_data['판매가']/1000, scatter_data['판매지수'], alpha=0.4, c='#667eea', s=30, edgecolors='white', linewidth=0.5)
z = np.polyfit(scatter_data['판매가']/1000, scatter_data['판매지수'], 1)
p = np.poly1d(z)
x_line = np.linspace(scatter_data['판매가'].min()/1000, scatter_data['판매가'].max()/1000, 100)
ax.plot(x_line, p(x_line), '--', color='#f5576c', linewidth=2, label=f'Trend (r={scatter_data[["판매가","판매지수"]].corr().iloc[0,1]:.3f})')
ax.set_xlabel('Price (KRW x1,000)', fontsize=11)
ax.set_ylabel('Sales Index', fontsize=11)
ax.set_title('Price vs Sales Index', fontsize=14, fontweight='bold', pad=15)
ax.legend(fontsize=10)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/07_price_vs_sales.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 7 done')

# 8. Yearly Avg Sales Index Trend
fig, ax = plt.subplots(figsize=(12, 6), facecolor=BG)
ax.set_facecolor(BG)
yp = df.groupby('출판일_연도')['판매지수'].mean().dropna()
recent_yp = yp[yp.index >= 2020]
ax.plot(recent_yp.index.astype(str), recent_yp.values, 'o-', color='#f5576c', linewidth=2.5, markersize=8, markerfacecolor='white', markeredgewidth=2)
for x, y in zip(recent_yp.index.astype(str), recent_yp.values):
    ax.text(x, y + 150, f'{int(y):,}', ha='center', fontsize=9, fontweight='bold')
ax.set_xlabel('Year', fontsize=11)
ax.set_ylabel('Avg Sales Index', fontsize=11)
ax.set_title('Average Sales Index by Publication Year (2020-2026)', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/08_yearly_avg_sales.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 8 done')

# 9. Discount Distribution
fig, ax = plt.subplots(figsize=(8, 6), facecolor=BG)
ax.set_facecolor(BG)
disc = df['할인율'].dropna().value_counts().sort_index()
bars = ax.bar(disc.index, disc.values, color=['#43e97b', '#f5576c', '#667eea'][:len(disc)], edgecolor='white', width=0.4)
for bar, val in zip(bars, disc.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5, f'{val}\n({val/len(df)*100:.1f}%)', ha='center', fontsize=10, fontweight='bold')
ax.set_xlabel('Discount Rate', fontsize=11)
ax.set_ylabel('Number of Books', fontsize=11)
ax.set_title('Discount Rate Distribution', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/09_discount_distribution.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 9 done')

# 10. KPI Summary Dashboard
fig, axes = plt.subplots(2, 3, figsize=(15, 8), facecolor=BG)
kpis = [
    ('Total Books', f'{len(df):,}', '#667eea'),
    ('Avg Price', f'{df["판매가"].mean():,.0f}KRW', '#764ba2'),
    ('Avg Sales Index', f'{df["판매지수"].mean():,.0f}', '#f5576c'),
    ('Publishers', f'{df["출판사"].nunique()}', '#4facfe'),
    ('Avg Discount', f'{df["할인율_num"].mean():.1f}%', '#43e97b'),
    ('Max Sales Index', f'{df["판매지수"].max():,.0f}', '#fa709a'),
]
for ax, (label, value, color) in zip(axes.flat, kpis):
    ax.set_facecolor(color)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.text(0.5, 0.65, value, ha='center', va='center', fontsize=22, fontweight='bold', color='white')
    ax.text(0.5, 0.25, label, ha='center', va='center', fontsize=12, color='white', alpha=0.9)
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
fig.suptitle('Key Performance Indicators', fontsize=16, fontweight='bold', y=0.98, color='#1a1a2e')
plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig(f'{OUT}/10_kpi_dashboard.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 10 done')

# 11. Top Authors
fig, ax = plt.subplots(figsize=(10, 6), facecolor=BG)
ax.set_facecolor(BG)
auth = df['저자'].value_counts().head(10).sort_values()
bars = ax.barh(range(10), auth.values, color=COLORS[:10], edgecolor='white', height=0.6)
ax.set_yticks(range(10))
ax.set_yticklabels(auth.index, fontsize=10)
for bar, val in zip(bars, auth.values):
    ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2, str(val), va='center', fontsize=10, fontweight='bold')
ax.set_xlabel('Number of Books', fontsize=11)
ax.set_title('Top 10 Most Prolific Authors', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/11_top_authors.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 11 done')

# 12. Yearly Price Trend
fig, ax = plt.subplots(figsize=(12, 6), facecolor=BG)
ax.set_facecolor(BG)
yp_price = df.groupby('출판일_연도')['판매가'].mean().dropna()
recent_pp = yp_price[yp_price.index >= 2020]
ax.plot(recent_pp.index.astype(str), recent_pp.values/1000, 's-', color='#764ba2', linewidth=2.5, markersize=8, markerfacecolor='white', markeredgewidth=2)
for x, y in zip(recent_pp.index.astype(str), recent_pp.values/1000):
    ax.text(x, y + 0.5, f'{y:.1f}K', ha='center', fontsize=9, fontweight='bold')
ax.set_xlabel('Year', fontsize=11)
ax.set_ylabel('Avg Price (KRW x1,000)', fontsize=11)
ax.set_title('Average Price Trend by Publication Year (2020-2026)', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(f'{OUT}/12_yearly_price_trend.png', dpi=150, bbox_inches='tight', facecolor=BG)
plt.close()
print('Chart 12 done')

print('\nAll charts saved to:', OUT)
