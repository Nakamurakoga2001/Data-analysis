import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os
from matplotlib import font_manager
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score

# Excelファイルからデータを読み込みます
file_path = 'エクセル分析用.xlsx'
data = pd.read_excel(file_path)
output_dir = '客観データ分析/graph年'
os.makedirs(output_dir, exist_ok=True)

# 日本語フォント（MSゴシック）の設定
font_path = '/Users/nakamurakouga/Library/Fonts/msgothic.ttc'
if os.path.exists(font_path):
    font_prop = font_manager.FontProperties(fname=font_path)
    plt.rcParams['font.sans-serif'] = [font_prop.get_name()]
    print(f"フォント設定完了: {font_prop.get_name()}")
else:
    print("エラー: フォントファイルが見つかりません。パスを確認してください。")
    
# 実際のカラム名に合わせて変更してください。
completion_year_column = '竣工年'  
casbee_wo_column = '合計点数'
rank_column = 'ランク'

plt.rcParams.update({
    'font.size': 14,         # 全体のフォントサイズ
    'axes.titlesize': 18,    # タイトルのフォントサイズ
    'axes.labelsize': 16,    # 軸ラベルのフォントサイズ
    'legend.fontsize': 14,   # 凡例のフォントサイズ
    'xtick.labelsize': 12,   # x軸目盛りのフォントサイズ
    'ytick.labelsize': 12    # y軸目盛りのフォントサイズ
})

# 散布図の作成
fig, ax1 = plt.subplots(figsize=(12, 5))

# ランクごとに異なる色とマーカーでプロット
for rank, marker, color in zip(['S', 'A', 'B+', 'B-'], ['o','^', 's', 'D'], ['#F8C3AA','#33A48F', '#457DBF', '#E87441']):
    rank_data = data[data[rank_column] == rank]
    ax1.scatter(rank_data[completion_year_column], rank_data[casbee_wo_column], 
                label=f'{rank}ランク', marker=marker, color=color)

# 軸ラベルと凡例の設定
ax1.set_xlabel('竣工年')
ax1.set_ylabel('CASBEE-不動産WO')
ax1.legend()

ax1.axvline(x=2006, color='#7DCACB', linestyle='-', linewidth=1, label='バリアフリー新法')

# グリッドを表示
ax1.grid(which='both', linestyle='--', linewidth=0.5)

# 回帰直線の計算
X = data[completion_year_column].values.reshape(-1, 1)  # 2次元配列に変換
y = data[casbee_wo_column].values

# 線形回帰モデルの作成と学習
model = LinearRegression()
model.fit(X, y)

# 回帰直線を描画
X_range = np.linspace(X.min(), X.max(), 100).reshape(-1, 1)
y_pred = model.predict(X_range)
ax1.plot(X_range, y_pred, color="#E6003A", linestyle="--", label="回帰直線")

# 決定係数 R^2 の計算
r2 = r2_score(y, model.predict(X))

# 式と R^2 をプロットに表示
equation_text = f'y = {model.coef_[0]:.2f} * x + {model.intercept_:.2f}\n$R^2$ = {r2:.2f}'
ax1.text(0.05, 0.95, equation_text, transform=ax1.transAxes, 
         fontsize=14, verticalalignment='top', bbox=dict(facecolor='white', alpha=0.5))

# グラフを表示
plt.show()

title = "竣工年とCASBEE-不動産WOの関係"
output_path = os.path.join(output_dir, f"{title}.png")
fig.savefig(output_path, format='png', dpi=300, bbox_inches='tight')
print(f"{title} のグラフを保存しました: {output_path}")
