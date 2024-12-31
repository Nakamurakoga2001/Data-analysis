import pandas as pd
'''
file_path = '回答者分析用.xlsx'
df1 = pd.read_excel(file_path, sheet_name='Sheet1')

print(df1.columns.to_list())

df = pd.DataFrame()

df['CASBEEフラグ'] = df1['合計点数'].apply(lambda x : 1 if x >= 50 else 0)

addlist = ['総合満足度','農業、林業フラグ', '漁業フラグ', '鉱業，採石業，砂利採取業フラグ', '建設業フラグ',
           '製造業フラグ', '電気・ガス・熱供給・水道業フラグ', '情報通信業フラグ', '運輸業、郵便業フラグ',
           '卸売業、小売業フラグ', '金融業、保険業フラグ', '不動産業，物品賃貸業フラグ', '学術研究，専門・技術サービス業フラグ',
           '宿泊業，飲食サービス業フラグ', '生活関連サービス業，娯楽業フラグ', '教育，学習支援業フラグ', '医療、福祉フラグ',
           '複合サービス業務フラグ', 'サービス業（他に分類されない）フラグ', '公務（他に分類されない）フラグ'
           ]

for column in addlist:
    if column in df1.columns:
     df[column] = df1[column]
     
print(df.columns.to_list())

output_file_path1 = '客観主観/二元配置分散分析/for2wayanova.xlsx'
df.to_excel(output_file_path1, index=False)
print(f"{output_file_path1}に保存済み ")

'''

import statsmodels.api as sm
from statsmodels.formula.api import ols
import seaborn as sns
import matplotlib.pyplot as plt
import os
from matplotlib import font_manager

font_path = '/Users/nakamurakouga/Library/Fonts/msgothic.ttc'
if os.path.exists(font_path):
    font_prop = font_manager.FontProperties(fname=font_path)
    plt.rcParams['font.sans-serif'] = [font_prop.get_name()]
    print(f"フォント設定完了: {font_prop.get_name()}")
else:
    print("エラー: フォントファイルが見つかりません。パスを確認してください。")
    
    
file2_path = 'エクセル分析用.xlsx'
df2 = pd.read_excel(file2_path, sheet_name='Sheet1')

df3 = pd.DataFrame()

df3['CASBEEフラグ'] = df2['合計点数'].apply(lambda x : 1 if x >= 50 else 0)
addlist2 = ['総合満足度','規模フラグ']
for column in addlist2:
    if column in df2.columns:
     df3[column] = df2[column]
     
print(df3.columns.to_list())

# 二元配置分散分析
df3['規模フラグ'] = df3['規模フラグ'].astype('category')
df3['CASBEEフラグ'] = df3['CASBEEフラグ'].astype('category')
model = ols('総合満足度 ~ C(CASBEEフラグ) * C(規模フラグ)', data=df3).fit()
anova_table = sm.stats.anova_lm(model, typ=2)
print("二元配置分散分析の結果:\n", anova_table)

# 交互作用を視覚化
fig, ax = plt.subplots(figsize=(10, 6))
sns.pointplot(data=df3, x='CASBEEフラグ', y='総合満足度', hue='規模フラグ', dodge=True, markers=["o", "s"], capsize=.1, ax=ax)

# グラフの装飾
plt.title("CASBEEフラグと規模フラグの交互作用による総合満足度の変化")
plt.xlabel("CASBEEフラグ")
plt.ylabel("総合満足度")
plt.legend(title="規模フラグ")
plt.grid(which='both', linestyle='--', linewidth=0.5)

# グラフの保存
output_dir = '客観主観/graph主観客観'
os.makedirs(output_dir, exist_ok=True)
title = "二元配置延べ"
output_path = os.path.join(output_dir, f"{title}.png")
fig.savefig(output_path, format='png', dpi=300, bbox_inches='tight')
print(f"{title} のグラフを保存しました: {output_path}")

# グラフを表示
plt.show()