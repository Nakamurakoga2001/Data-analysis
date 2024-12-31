import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os
from matplotlib import font_manager

# 日本語フォント（MSゴシック）の設定
font_path = '/Users/nakamurakouga/Library/Fonts/msgothic.ttc'
if os.path.exists(font_path):
    font_prop = font_manager.FontProperties(fname=font_path)
    plt.rcParams['font.sans-serif'] = [font_prop.get_name()]
    print(f"フォント設定完了: {font_prop.get_name()}")
else:
    print("エラー: フォントファイルが見つかりません。パスを確認してください。")

# データの読み込み
file_path = '分析用.xlsx'
df = pd.read_excel(file_path)

# 保存先ディレクトリの指定
output_dir = '客観データ分析/graph'
os.makedirs(output_dir, exist_ok=True)

# カテゴリ別にコラムリストを設定
columns_set1 = [
    '1.1躯体の耐震性能', '1.2設備の信頼性', '1.3災害時エネルギー供給', '1.4自然災害リスク対策', '1.5BCPの有無',
    '2.1セキュリティ設備', '2.2バリアフリー法への対応', '2.3土壌環境品質・ブラウンフィールド再生', '1.1外観デザイン',
    '2.1オフィスからの眺望', '2.2生物多様性の向上', '2.3トイレの充足性・機能性', '2.4リフレッシュスペース',
    '3.1建築物衛生基準への適合状況', '3.2自然換気性能', '3.3自然光の導入', '3.4分煙対応、禁煙対応', '4.1維持管理',
    '4.2満足度調査の定期的実施等', '4.3健康維持・増進プログラム', '1.1空間の形状・自由さ',
    '1.2動線における出会いの場の創出', '1.3打ち合わせスペース', '2.1高度情報通信インフラ', '2.2情報共有インフラ'
]

columns_set2 = [
    '防災対策', '安心安全対策', 'デザイン性', 'リフレッシュ_y', '室内環境質', '維持管理・運営', '空間・内装',
    '情報通信'
]

columns_set3 = ['Qw1安心・安全', 'Qw2健康性・快適性', 'Qw3知的生産性向上']
column_mean = '平均点'
column_total = '合計点数'
column_rank = 'ランク'

# カラーと斜線パターンの指定
colors = ["#3080C4", "#ED7851", "#40B6AA", "#4D959B"]
hatch_patterns = ['//////', '------', '\\\\\\']

# ヒストグラムを描画してPNGとして保存する関数
def plot_and_save_histogram(data, title, xlabel, bins, color, hatch, output_path=None):
    fig, ax = plt.subplots(figsize=(5, 3))
    hist_data, bins, patches = ax.hist(data, bins=bins, color=color, edgecolor=color, alpha=0.7)
    
    # 各棒に斜線パターンを設定
    for patch in patches:
        patch.set_hatch(hatch)

    # 統計情報の計算
    mean_val = data.mean()
    median_val = data.median()
    std_dev = data.std()
    sample_count = len(data)

    # 統計情報をグラフ上に表示
    stats_text = (
        f"平均: {mean_val:.2f}\n"
        f"中央値: {median_val:.2f}\n"
        f"標準偏差: {std_dev:.2f}\n"
        f"サンプル数: {sample_count}"
    )
    ax.text(0.95, 0.95, stats_text, transform=ax.transAxes, fontsize=8,
            verticalalignment='top', horizontalalignment='right', bbox=dict(facecolor='white', alpha=0.5))

    # ラベルの設定
    ax.set_title(title, fontsize=10)
    ax.set_xlabel(xlabel, fontsize=8)
    ax.set_ylabel('頻度', fontsize=8)
    
    # 縦軸線と横軸線を非表示
    # ax.spines['top'].set_visible(False)
    # ax.spines['right'].set_visible(False)
    # ax.spines['left'].set_visible(False)
    # ax.spines['bottom'].set_visible(False)
    ax.grid(False)
    ax.set_xticks(bins[:-1] + (bins[1] - bins[0]) / 2)
    ax.set_xticklabels([f"{b:.1f}" if b % 1 != 0 else f"{int(b)}" for b in bins[:-1]])

    # 合計点数ヒストグラムの場合、5点刻みで表示
    if title == "合計点数":
        ax.set_xticks(np.arange(25, 86, 5))
        ax.set_xticklabels([f"{int(b)}" for b in np.arange(25, 86, 5)])

    if output_path:
        fig.savefig(output_path, format='png', dpi=300, bbox_inches='tight')
        print(f"{title} のグラフを保存しました: {output_path}")
    plt.close(fig)
    
# ヒストグラムの設定に基づいて描画と保存
for idx, col in enumerate(columns_set1):
    color = colors[0]
    hatch = hatch_patterns[0]
    output_path = os.path.join(output_dir, f"{col}_histogram.png")
    plot_and_save_histogram(
        df[col].dropna(), 
        f'{col}', 
        'スコア', 
        bins=[1, 2, 3, 4, 5, 6],  # 1～5のスコア
        color=color, 
        hatch=hatch,
        output_path=output_path
    )

# カテゴリに基づいたヒストグラム（別のビン設定）
for idx, col in enumerate(columns_set2):
    color = colors[1]
    hatch = hatch_patterns[1]
    output_path = os.path.join(output_dir, f"{col}_histogram.png")
    plot_and_save_histogram(
        df[col].dropna(), 
        f'{col}', 
        'スコア', 
        bins=np.arange(0, 6, 0.5),  # データに応じて10個のビンで調整
        color=color, 
        hatch=hatch,
        output_path=output_path
    )

# Qw列に基づいたヒストグラム（1～5のビン）
for idx, col in enumerate(columns_set3):
    color = colors[2]
    hatch = hatch_patterns[2]
    output_path = os.path.join(output_dir, f"{col}_histogram.png")
    plot_and_save_histogram(
        df[col].dropna(), 
        f'{col}', 
        'スコア', 
        bins=np.arange(0, 6, 0.5),  # 1～5のスコア
        color=color, 
        hatch=hatch,
        output_path=output_path
    )

# 平均点のヒストグラム（0～100の範囲、10点刻みのビン）
output_path = os.path.join(output_dir, "平均点_histogram.png")
plot_and_save_histogram(
    df[column_mean].dropna(), 
    '平均点', 
    '平均点', 
    bins=np.arange(0, 6, 0.5),  # 0～100、10点刻み
    color=colors[0], 
    hatch=hatch_patterns[0],
    output_path=output_path
)

# 合計点数のヒストグラム（25から85の範囲、1点刻みのビン）
output_path = os.path.join(output_dir, "合計点数_histogram.png")
plot_and_save_histogram(
    df[column_total].dropna(), 
    '合計点数', 
    '合計点数', 
    bins=np.arange(19, 86, 5),  # 25～85、1点刻み
    color="#4D959B", 
    hatch=hatch_patterns[1],
    output_path=output_path,
    
)

# ランクの分布を棒グラフで追加
fig, ax = plt.subplots(figsize=(5, 3))
rank_order = ['C', 'B-', 'B+', 'A', 'S']
df[column_rank].value_counts().reindex(rank_order, fill_value=0).plot(kind='bar', ax=ax, edgecolor=colors[2], color=colors[2], hatch=hatch_patterns[2], alpha=0.7)
ax.set_title('ランクの分布', fontsize=10)
ax.set_xlabel('ランク', fontsize=8)
ax.set_ylabel('頻度', fontsize=8)
output_path = os.path.join(output_dir, "ランク_distribution.png")

# ランクの分布の縦軸線と横軸線を非表示
# ax.spines['top'].set_visible(False)
# ax.spines['right'].set_visible(False)
# ax.spines['left'].set_visible(False)
# ax.spines['bottom'].set_visible(False)
ax.grid(False)

# ランクの棒グラフを保存
fig.savefig(output_path, format='png', dpi=300, bbox_inches='tight')
print("ランクの分布の棒グラフが追加されました")
plt.close(fig)
