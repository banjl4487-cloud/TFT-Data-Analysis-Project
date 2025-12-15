import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.ticker import FuncFormatter
import warnings
import os

warnings.filterwarnings("ignore", message="Glyph .* missing from font")

# --- 데이터 불러오기 및 전처리 (기존과 동일) ---
file_path = 'vi_top10_items_with_summary.xlsx'
df = pd.read_excel(file_path)

def clean_text_from_data(text):
    if isinstance(text, str):
        return text.replace('■', '').replace('□', '')
    return text
df['아이템 이름'] = df['아이템 이름'].apply(clean_text_from_data)

# --- ★★★★★★★★★★ 영문 전용 폰트 설정 (한글 폰트 관련 설정 싹 다 제거!) ★★★★★★★★★★ ---

# Matplotlib의 폰트 캐시는 초기화 한 번 해주면 좋음 (터미널에서 수동으로!)
# python -c "import matplotlib; import matplotlib.font_manager as fm; fm._clear_cached_fonts(); print('Cache cleared. Restart IDE.')"

plt.rcParams['font.family'] = 'sans-serif'  # 영문 전용 기본 고딕 폰트
plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans', 'sans-serif'] # 가장 범용적인 영문 폰트
plt.rcParams['axes.unicode_minus'] = False # 마이너스 기호 깨짐 방지 (이건 그대로!)
plt.rcParams['text.usetex'] = False # TeX 시스템 사용 안 함

# --- ★★★★★★★★★★ 폰트 설정 끝 ★★★★★★★★★★ ---


# 시각화 기본 설정
sns.set_style("whitegrid")
plt.figure(figsize=(22, 16))

# 데이터 정렬 및 라벨링
df_sorted = df.sort_values(by='장착 횟수', ascending=False).reset_index(drop=True)
# 아이템 이름도 영문으로 바꿔줘야 깔끔하겠지? (필요하면 df['아이템 이름']을 영문으로 매핑하는 코드 추가)
df_sorted['순위_아이템_표시'] = (df_sorted.index + 1).astype(str) + ". " + df_sorted['아이템 이름']

# 바 그래프 그리기 (방어아이템 여부별 색 구분)
bar_plot = sns.barplot(
    x='순위_아이템_표시',
    y='장착 횟수',
    hue='방어 아이템 여부',
    data=df_sorted,
    palette={'방어아이템': '#1f77b4', '비 방어아이템': '#ff7f0e'}
)

# 제목 및 축 레이블 (전부 영문으로 변경!)
plt.title('VI Champion Item Visualization (Defense Items)', fontweight='bold', fontsize=40)
plt.ylabel('Equip Count', rotation=0, ha='right', va='center', labelpad=40, fontsize=30)


# X축 눈금 라벨 회전 및 크기
plt.xticks(rotation=45, ha='right', fontsize=25)
plt.yticks(fontsize=25)

# Y축 콤마 포매터 적용
def comma_formatter(x, pos):
    return f'{int(x):,}'
plt.gca().yaxis.set_major_formatter(FuncFormatter(comma_formatter))

# 막대 위 값 표시
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

# 범례 설정 (영문으로 변경!)
plt.legend(title='Item Type', title_fontsize=0, loc='upper right', labels=['Defense Item', 'Non-Defense Item'], fontsize=30)

plt.tight_layout()
plt.show()