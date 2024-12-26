import pandas as pd

file_path = '客観データ/物件リスト.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

first_colums = ['物件名','物件CD','WO','RE','バリアフリー法','建築物衛生法']
survey_columns = [
    '防災', '防犯', '感染症対策', 'ダイバーシティ', '光', '温熱',
    '空気', '音', '清掃', '緑化', 'トイレ', 'エレベーター', 
    '健康施策', '内装', 'ワークスペース', 'リフレッシュ',
    'コミュニケーション', '通信', '景観', '利便性', '地域愛着','リレーション'
]

df['合計点'] = df[survey_columns].sum(axis=1)
df['竣工年'] = pd.to_datetime(df['竣工日']).dt.year

df['規模(大/中小)'] = df['規模'].apply(lambda x: 1 if x == '大規模' else 0)
df['施設分類'] = df['施設分類'].apply(lambda x: 1 if x == '事務所' else 0)

syuto_area = ['東京都', '神奈川県', '千葉県', '埼玉県']
area1_flags = []  
for i in df['都道府県']:
    if i in syuto_area:
        area1_flags.append(1)
    else:
        area1_flags.append(0)
df['首都圏フラグ'] = area1_flags

sanndaitoshi_area = ['東京都', '大阪府', '愛知県']
area2_flags = []  
for i in df['都道府県']:
    if i in sanndaitoshi_area:
        area2_flags.append(1)
    else:
        area2_flags.append(0)
df['三大都市圏フラグ'] = area2_flags
# df['三大都市圏フラグ'] = df['都道府県'].apply(lambda x: 1 if x in sanndaitoshi_area else 0)


# 自分の書き方が悪いせいで、このコメントアウトを外すと、全部実行し直す必要があります。
# その勇気はあなたにはありますか？別にここで外さないほうがいいからな
# 業種分類の19種類をリスト化
industry_categories = [
    '農業、林業', '漁業', '鉱業，採石業，砂利採取業', '建設業', '製造業',
    '電気・ガス・熱供給・水道業', '情報通信業', '運輸業、郵便業', '卸売業、小売業',
    '金融業、保険業', '不動産業，物品賃貸業', '学術研究，専門・技術サービス業',
    '宿泊業，飲食サービス業', '生活関連サービス業，娯楽業', '教育，学習支援業',
    '医療、福祉', '複合サービス業務', 'サービス業（他に分類されない）', '公務（他に分類されない）'
]

# 業種ごとにフラグを立てる
for category in industry_categories:
    # 各業種のフラグを立てる列を追加（業種が一致すれば1、そうでなければ0）
    df[f'{category}フラグ'] = df['業種分類'].apply(lambda x: 1 if x == category else 0)


industry_flags = df[f'{industry_categories}フラグ']



analysis_kyakkann_colums = (
    first_colums + 
    ['竣工年', '延床（㎡）', '規模(大/中小)', '首都圏フラグ', '三大都市圏フラグ',] +
    survey_columns +['合計点']+  ['総合満足度'] 
)
   
   
print(df.columns)
print(df[analysis_kyakkann_colums].head(10))

output_file_path = '客観データ/for_anarlysis_物件属性業種あり.xlsx'
df[analysis_kyakkann_colums].to_excel(output_file_path, index=False)
print(f"データのクリーニングお疲れさまです。 {output_file_path}に保存できましたよ ")