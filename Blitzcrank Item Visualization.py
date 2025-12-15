import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.ticker import FuncFormatter
import warnings
import os

warnings.filterwarnings("ignore", message="Glyph .* missing from font")

# --- 데이터 불러오기 ---
file_path = 'blitzcrank_top10_items_with_summary.xlsx'  # 블리츠 데이터 파일명
df = pd.read_excel(file_path)

# --- 텍스트 클린징 (■, □ 제거) ---
def clean_text_from_data(text):
    if isinstance(text, str):
        return text.replace('■', '').replace('□', '')
    return text
df['아이템 이름'] = df['아이템 이름'].apply(clean_text_from_data)

# --- 폰트 설정: 영문 전용 범용 sans-serif 계열 ---
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False # 마이너스 기호 깨짐 방지
plt.rcParams['text.usetex'] = False # TeX 시스템 사용 안 함


# --- 시각화 환경 설정 ---
sns.set_style("whitegrid")
plt.figure(figsize=(22, 16))

# --- 데이터 정렬 및 X축 라벨링 (순위 + 아이템 이름!) ---
df_sorted = df.sort_values(by='장착 횟수', ascending=False).reset_index(drop=True)
# 이제 X축에 '순위. 아이템 이름' 모두 표시된다!
df_sorted['아이템_랭크_이름'] = (df_sorted.index + 1).astype(str) + ". " + df_sorted['아이템 이름']

# --- 막대 그래프 그리기 ---
bar_plot = sns.barplot(
    x='아이템_랭크_이름',  # 순위 + 아이템 이름 사용!
    y='장착 횟수',
    hue='방어 아이템 여부',  # 방어아이템 / 비방어아이템 구분
    data=df_sorted,
    palette={'방어아이템': '#1f77b4', '비 방어아이템': '#ff7f0e'}
)

# --- 제목 및 Y축 라벨 ---
plt.title('Blitzcrank Item Visualization (Defense Items)', fontweight='bold', fontsize=40)
plt.ylabel('Equip Count', rotation=0, ha='right', va='center', labelpad=40, fontsize=30)

# --- X축 축명 제거 (여전히 빈 문자열로 깔끔하게!) ---
plt.xlabel('')

# --- X축 눈금 라벨 (아이템 이름+순위) 45도 회전 및 크기 조절! ---
plt.xticks(rotation=45, ha='right', fontsize=25) # 45도 회전! 폰트 크기 20으로 시원하게!
plt.yticks(fontsize=25) # Y축 눈금 폰트 크기

# --- Y축 콤마 포맷터 ---
def comma_formatter(x, pos):
    return f'{int(x):,}'
plt.gca().yaxis.set_major_formatter(FuncFormatter(comma_formatter))

# --- 막대 위 착용 횟수 표시 ---
max_count = df_sorted['장착 횟수'].max()
for p in bar_plot.patches:
    height = p.get_height()
    if height == 0:
        continue
    label_text = f'{int(height):,}'
    bar_plot.text(
        p.get_x() + p.get_width() / 2., height + max_count * 0.01,
        label_text, ha='center', color='black', weight='bold', fontsize=25
    )

# --- 범례 설정 (영어 labels!) ---
plt.legend(
    title='Item Type',
    loc='upper right',
    labels=['Defense Item', 'Non-Defense Item'], # 범례 텍스트 여전히 영어!
    fontsize=30
)

plt.tight_layout()
plt.show()