import pandas as pd

infofile_path = '主観データ/主観分析用.xlsx'
df1 = pd.read_excel(infofile_path, sheet_name='Sheet1')
print('hitotumeyomikomiok')
evalfile_path = '客観データ/集計行列_20241104_1638.xlsx'##これを変えろよ！！
#ヒルコート東新宿の名前がおかしくなってるから、ヒルコート東新宿にかえてね。新しくしゅうけいするなら！！117585がぶっけんCD 
df2 = pd.read_excel(evalfile_path, sheet_name='物件データ')
print('hutatumeyomikomiok')

# 業種フラグ列の確認
industry_columns = [col for col in df1.columns if 'フラグ' in col]
print("業種フラグ列:", industry_columns)

# カテゴリごとの評価項目と平均値を算出
s_qw1_cols = ['防災', '防犯', 'ダイバーシティ']
s_qw2_cols = ['光', '温熱', '空気', '音', '清掃', '緑化', 'トイレ', 'エレベーター', '健康施策', '感染症対策', 'リフレッシュ', '景観']
s_qw3_cols = ['内装', 'ワークスペース', 'コミュニケーション', '通信', 'リレーション', '利便性', '地域愛着']
df1['S安心・安全'] = df1[s_qw1_cols].mean(axis=1)
df1['S健康性・快適性'] = df1[s_qw2_cols].mean(axis=1)
df1['S知的生産性向上'] = df1[s_qw3_cols].mean(axis=1)

# データフレームのキー列の前処理
df1['物件CD'] = df1['物件CD'].astype(str).str.strip()
df2['物件CD'] = df2['物件CD'].astype(str).str.strip()

# データフレームの結合
merged_df = pd.merge(df1, df2, on='物件CD', how='inner')
print("結合後のデータフレーム列:", merged_df.columns)

rank_mapping = {'B-': 0, 'B+': 1, 'A': 2, 'S': 3}
merged_df['ランクフラグ'] = merged_df['ランク'].map(rank_mapping)

# # 業種フラグが結合後に含まれているかの確認
# merged_industry_columns = [col for col in merged_df.columns if 'フラグ' in col]
# print("結合後のフラグ列:", merged_industry_columns)

# # 列名を短縮するためのマッピング
# short_col_names = {col: col.replace('業、', '').replace('、', '').replace('フラグ', '') for col in merged_industry_columns}
# merged_df.rename(columns=short_col_names, inplace=True)
# データの出力
output_file_path = '/Users/nakamurakouga/Documents/Chiba-U/2024研究関連/1ザイマックス/分析用/回答者分析用.xlsx'
merged_df.to_excel(output_file_path, index=False)
print(f"データの結合お疲れさまです。{output_file_path} に保存できました。")
