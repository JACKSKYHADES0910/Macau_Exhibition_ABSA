import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from wordcloud import WordCloud
import jieba
import os

# 配置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 150

# 检查可用的字体路径供词云使用
font_paths = ['C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf', 'C:/Windows/Fonts/STZHONGS.TTF']
font_path = next((f for f in font_paths if os.path.exists(f)), None)

OUTPUT_DIR = 'd:/Project/Macau_Sentiment_Analysis/'

print("Loading data...")
# 优先读取最终版 xlsx，其次读取中间态 csv
xlsx_path = os.path.join(OUTPUT_DIR, '澳门展览_清洗分析版_v2.xlsx')
csv_path = os.path.join(OUTPUT_DIR, '澳门展览_情感分析_中间态_v2.csv')
if os.path.exists(xlsx_path):
    df = pd.read_excel(xlsx_path)
    print(f"Loaded from {xlsx_path} ({len(df)} records)")
elif os.path.exists(csv_path):
    df = pd.read_csv(csv_path, encoding='utf-8-sig')
    print(f"Loaded from {csv_path} ({len(df)} records)")
else:
    print("No data file found!")
    exit()

# 过滤掉解析失败的行
dimensions = [
    "文化认知价值",
    "体验审美价值",
    "伦理公共价值",
    "象征价值",
    "感知文化真实性",
    "忠诚",
    "声誉"
]

# 学术配色方案
COLORS = {
    'primary': '#2563EB',
    'secondary': '#10B981',
    'accent': '#F59E0B',
    'danger': '#EF4444',
    'gray': '#D1D5DB',
    'dark': '#1F2937',
    'palette': ['#2563EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#06B6D4']
}


# ===== 图表一：七维提及率雷达图 =====
def draw_radar_chart():
    print("  [1/6] Generating Radar Chart...")
    mention_rates = {}
    for dim in dimensions:
        col = f"{dim}_情感极性"
        if col in df.columns:
            valid = df[col].dropna()
            valid = valid[~valid.isin(['解析失败'])]
            if len(valid) > 0:
                mention_rates[dim] = (valid == '提及').sum() / len(valid) * 100
            else:
                mention_rates[dim] = 0

    labels = list(mention_rates.keys())
    stats = list(mention_rates.values())
    
    angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False)
    stats_closed = np.concatenate((stats, [stats[0]]))
    angles_closed = np.concatenate((angles, [angles[0]]))
    
    fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True))
    ax.plot(angles_closed, stats_closed, 'o-', linewidth=2.5, color=COLORS['primary'], markersize=8)
    ax.fill(angles_closed, stats_closed, color=COLORS['primary'], alpha=0.15)
    
    # 在每个节点标注数值
    for angle, stat, label in zip(angles, stats, labels):
        ax.text(angle, stat + 4, f'{stat:.1f}%', ha='center', va='bottom', fontsize=10, fontweight='bold', color=COLORS['dark'])
    
    ax.set_thetagrids(angles * 180/np.pi, labels, fontsize=12)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(['20%', '40%', '60%', '80%', '100%'], fontsize=8, color='gray')
    ax.set_title("各维度提及率(%)雷达图", weight='bold', size=16, pad=30)
    ax.grid(color='#E5E7EB', linewidth=0.5)
    
    plt.savefig(OUTPUT_DIR + '1_七维提及率雷达图.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()


# ===== 图表二：各维度提及频次排名柱状图 =====
def draw_ranking_bar():
    print("  [2/6] Generating Ranking Bar Chart...")
    mention_data = []
    for dim in dimensions:
        col = f"{dim}_情感极性"
        if col in df.columns:
            valid = df[col].dropna()
            valid = valid[~valid.isin(['解析失败'])]
            mention_count = (valid == '提及').sum()
            total = len(valid)
            mention_data.append({
                '维度': dim,
                '提及数': mention_count,
                '总数': total,
                '提及率': mention_count / total * 100 if total > 0 else 0
            })
    
    df_rank = pd.DataFrame(mention_data).sort_values('提及率', ascending=True)
    
    fig, ax = plt.subplots(figsize=(10, 7))
    bars = ax.barh(df_rank['维度'], df_rank['提及率'], color=COLORS['palette'][:len(df_rank)], height=0.6, edgecolor='white', linewidth=0.5)
    
    # 在柱子末端标注数值
    for bar, row in zip(bars, df_rank.itertuples()):
        width = bar.get_width()
        ax.text(width + 1, bar.get_y() + bar.get_height()/2, f'{width:.1f}% ({row.提及数}/{row.总数})', va='center', fontsize=10, color=COLORS['dark'])
    
    ax.set_xlabel('提及率 (%)', fontsize=12)
    ax.set_xlim(0, max(df_rank['提及率']) + 15)
    ax.set_title('各维度提及率排名', weight='bold', size=16, pad=15)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(axis='y', labelsize=11)
    
    plt.savefig(OUTPUT_DIR + '2_各维度提及率排名.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()


# ===== 图表三：提及/未提及堆叠柱状图 (百分比) =====
def draw_stacked_bar():
    print("  [3/6] Generating Stacked Bar Chart...")
    data = []
    for dim in dimensions:
        col = f"{dim}_情感极性"
        if col in df.columns:
            valid = df[col].dropna()
            valid = valid[~valid.isin(['解析失败'])]
            counts = valid.value_counts()
            data.append({
                '维度': dim,
                '提及': counts.get('提及', 0),
                '未提及': counts.get('未提及', 0)
            })
            
    df_stack = pd.DataFrame(data).set_index('维度')
    df_pct = df_stack.div(df_stack.sum(axis=1), axis=0) * 100
    df_pct = df_pct.fillna(0)
    
    colors = [COLORS['secondary'], COLORS['gray']]
    
    fig, ax = plt.subplots(figsize=(12, 7))
    df_pct[['提及', '未提及']].plot(kind='barh', stacked=True, color=colors, ax=ax, edgecolor='white', linewidth=0.5)
    
    for c in ax.containers:
        labels = [f'{v.get_width():.1f}%' if v.get_width() > 5 else '' for v in c]
        ax.bar_label(c, labels=labels, label_type='center', fontsize=10, color='white', weight='bold')
        
    ax.set_xlabel('百分比 (%)', fontsize=12)
    ax.set_ylabel('')
    ax.set_title('各维度提及/未提及占比', weight='bold', size=16, pad=15)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.08), ncol=2, fontsize=11, frameon=False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(axis='y', labelsize=11)
    
    plt.savefig(OUTPUT_DIR + '3_提及占比堆叠柱状图.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()


# ===== 图表四：核心关注点词云图 =====
def draw_wordcloud():
    print("  [4/6] Generating Word Cloud...")
    if not font_path:
        print("    Warning: Cannot find valid Chinese font for wordcloud, skipping.")
        return
        
    mentioned_texts = []
    for dim in dimensions:
        polarity_col = f"{dim}_情感极性"
        reason_col = f"{dim}_判定原因提取"
        if polarity_col in df.columns and reason_col in df.columns:
            mentions = df[df[polarity_col] == '提及'][reason_col].dropna().tolist()
            mentioned_texts.extend([str(t) for t in mentions if str(t) not in ['无', '解析失败', 'nan']])
            
    full_text = " ".join(mentioned_texts)
    
    if not full_text.strip():
        print("    No valid text for wordcloud.")
        return
        
    words = jieba.lcut(full_text)
    stopwords = set(["的", "了", "是", "我", "在", "和", "就", "都", "不", "也", "有", "很",
                     "这个", "感觉", "觉得", "这", "没", "无", "因为", "还是", "但是", "但",
                     "比较", "非常", "可以", "一个", "到", "去", "看", "说", "让", "会",
                     "什么", "那", "被", "从", "把", "对", "做", "给", "用", "着"])
    filtered_words = [w for w in words if len(w) > 1 and w not in stopwords]
    text_for_wc = " ".join(filtered_words)
    
    if not text_for_wc.strip():
        print("    No valid words after filtering.")
        return
        
    wc = WordCloud(
        font_path=font_path,
        width=1200, height=700,
        background_color='white',
        colormap='viridis',
        max_words=200,
        random_state=42,
        margin=10,
        prefer_horizontal=0.7
    )
    wc.generate(text_for_wc)
    
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.imshow(wc, interpolation='bilinear')
    ax.axis('off')
    ax.set_title('游客核心关注点词云图', weight='bold', size=16, pad=15)
    
    plt.savefig(OUTPUT_DIR + '4_核心关注点词云图.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()


# ===== 图表五：各维度提及率对比甜甜圈组图 =====
def draw_donut_grid():
    print("  [5/6] Generating Donut Grid Chart...")
    fig, axes = plt.subplots(2, 4, figsize=(18, 9))
    axes = axes.flatten()
    
    for i, dim in enumerate(dimensions):
        ax = axes[i]
        col = f"{dim}_情感极性"
        if col in df.columns:
            valid = df[col].dropna()
            valid = valid[~valid.isin(['解析失败'])]
            mention = (valid == '提及').sum()
            not_mention = (valid == '未提及').sum()
            total = mention + not_mention
            rate = mention / total * 100 if total > 0 else 0
            
            colors = [COLORS['palette'][i], '#F3F4F6']
            wedges, _, autotexts = ax.pie(
                [mention, not_mention],
                autopct=lambda pct: f'{pct:.1f}%' if pct > 5 else '',
                startangle=90,
                colors=colors,
                wedgeprops=dict(width=0.35, edgecolor='white', linewidth=2)
            )
            plt.setp(autotexts, size=9, weight="bold")
            ax.text(0, 0, f'{rate:.0f}%', ha='center', va='center', fontsize=16, fontweight='bold', color=COLORS['palette'][i])
        
        ax.set_title(dim, fontsize=11, fontweight='bold', pad=8)
    
    # 隐藏第8个空位
    axes[7].axis('off')
    
    fig.suptitle('各维度提及率环形图对比', fontsize=18, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR + '5_各维度提及率环形图对比.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()


# ===== 图表六：数据摘要统计表 =====
def draw_summary_table():
    print("  [6/6] Generating Summary Table...")
    summary_data = []
    for dim in dimensions:
        col = f"{dim}_情感极性"
        if col in df.columns:
            valid = df[col].dropna()
            valid = valid[~valid.isin(['解析失败'])]
            mention = (valid == '提及').sum()
            not_mention = (valid == '未提及').sum()
            total = mention + not_mention
            rate = mention / total * 100 if total > 0 else 0
            summary_data.append([dim, total, mention, not_mention, f'{rate:.1f}%'])
    
    # 总计行
    total_records = sum(r[1] for r in summary_data)
    total_mention = sum(r[2] for r in summary_data)
    total_not = sum(r[3] for r in summary_data)
    avg_rate = total_mention / total_records * 100 if total_records > 0 else 0
    summary_data.append(['合计/均值', total_records, total_mention, total_not, f'{avg_rate:.1f}%'])
    
    col_labels = ['维度', '有效样本', '提及数', '未提及数', '提及率']
    
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.axis('off')
    
    table = ax.table(
        cellText=summary_data,
        colLabels=col_labels,
        cellLoc='center',
        loc='center',
        colWidths=[0.25, 0.15, 0.15, 0.15, 0.15]
    )
    
    # 样式
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1.0, 1.6)
    
    # 表头样式
    for j in range(len(col_labels)):
        cell = table[0, j]
        cell.set_facecolor(COLORS['primary'])
        cell.set_text_props(color='white', fontweight='bold')
    
    # 最后一行（合计行）加粗
    last_row = len(summary_data)
    for j in range(len(col_labels)):
        cell = table[last_row, j]
        cell.set_facecolor('#EFF6FF')
        cell.set_text_props(fontweight='bold')
    
    # 交替行背景
    for i in range(1, len(summary_data)):
        for j in range(len(col_labels)):
            if i % 2 == 0:
                table[i, j].set_facecolor('#F9FAFB')
    
    ax.set_title('各维度情感分析统计汇总表', weight='bold', size=16, pad=20)
    
    plt.savefig(OUTPUT_DIR + '6_统计汇总表.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()


if __name__ == "__main__":
    print(f"\nGenerating 6 charts for {len(df)} records...\n")
    draw_radar_chart()
    draw_ranking_bar()
    draw_stacked_bar()
    draw_wordcloud()
    draw_donut_grid()
    draw_summary_table()
    print(f"\nAll 6 charts saved to {OUTPUT_DIR}")
