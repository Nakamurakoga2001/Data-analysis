import pandas as pd

infofile_path = '客観データ/for_anarlysis_物件属性.xlsx'
df1 = pd.read_excel(infofile_path, sheet_name='Sheet1')
print('hitotumeyomikomiok')
evalfile_path = '客観データ/集計行列_20241104_1638.xlsx'##これを変えろよ！！
#ヒルコート東新宿の名前がおかしくなってるから、ヒルコート東新宿にかえてね。新しくしゅうけいするなら！！117585がぶっけんCD 
df2 = pd.read_excel(evalfile_path, sheet_name='物件データ')
print('hutatumeyomikomiok')
print(df1.columns)
print(df2.columns)

'''
分析用に主観クリーニングと評価項目の追加を行います。
'''
# カテゴリごとの評価項目と平均値を算出（主観と客観の整合のため）
s_qw1_cols = ['防災', '防犯', 'ダイバーシティ']
s_qw2_cols = ['光', '温熱', '空気', '音', '清掃', '緑化', 'トイレ', 'エレベーター', '健康施策', '感染症対策', 'リフレッシュ', '景観']
s_qw3_cols = ['内装', 'ワークスペース', 'コミュニケーション', '通信', 'リレーション', '利便性', '地域愛着']
df1['S安心・安全'] = df1[s_qw1_cols].mean(axis=1)
df1['S健康性・快適性'] = df1[s_qw2_cols].mean(axis=1)
df1['S知的生産性向上'] = df1[s_qw3_cols].mean(axis=1)
print(df1[['S安心・安全', 'S健康性・快適性', 'S知的生産性向上']].head())

# 初めて結合するときは、長さと列がちゃんと結合できているかしらべておこうね・
merged_df = pd.merge(df1, df2, on ='物件名', how = 'inner')
list1 = merged_df['物件名'].tolist()
count = len(merged_df['物件名'])
print(merged_df.columns)
print(f"物件名のリスト: {list1}")
print(f"物件名の総数: {count}")

#なんかだめらしい警告が出る。merged_df['ランクフラグ'] = merged_df['ランク'].replace({'B-': 0, 'B+': 1, 'A': 3, 'S': 4})
#cのぶっけんがあったら、ここは直しておけよ。
# ランクフラグを map を使って作成
rank_mapping = {'B-': 0, 'B+': 1, 'A': 2, 'S': 3}
merged_df['ランクフラグ'] = merged_df['ランク'].map(rank_mapping)

print(merged_df.columns)



output_file_path1 = '分析用.xlsx'
#output_file_path2 = ''
merged_df.to_excel(output_file_path1, index=False)
#merged_df.to_excel(output_file_path2, index=False)
print(f"データの結合お疲れさまです。 {output_file_path1}に保存できましたよ ")

""" 各種関係性まとめ
    L列から主観
    AI列から客観
    防災←→"1.5BCPの有無"
    防犯←→"2.1セキュリティ設備"
    ダイバーシティ←→"2.2バリアフリー法への対応"
    光←→"3.3自然光の導入"
    温熱←→"3.1建築物衛生基準への適合状況"
    空気←→"3.1建築物衛生基準への適合状況", "3.2自然換気性能"
    音←→
    清掃←→"4.1維持管理"
    緑化←→"2.2生物多様性の向上"
    トイレ←→"2.3トイレの充足性・機能性"
    エレベーター←→
    健康施策←→"4.3健康維持・増進プログラム"
    内装←→
    ワークスペース←→"1.1空間の形状・自由さ","1.3打ち合わせスペース"
    リフレッシュ←→"2.4リフレッシュスペース"
    コミュニケーション←→
    通信←→"2.1高度情報通信インフラ"
    景観←→"1.1外観デザイン"
    利便性←→
    地域愛着←→ "4.3健康維持・増進プログラム"
    リレーション←→"2.2情報共有インフラ"
    
    "S安心・安全" :['防災', '防犯', 'ダイバーシティ']
    "S健康性・快適性" :['光', '温熱', '空気', '音', '清掃', '緑化', 'トイレ', 'エレベーター', '健康施策', '感染症対策','リフレッシュ_x', '景観']
    "S知的生産性向上" :['内装', 'ワークスペース', 'コミュニケーション', '通信', 'リレーション', '利便性', '地域愛着']
    
    利便性？？地域愛着→一旦QW3
    """

