import pandas as pd
import os

# 新しいファイルの読み込み
file_path = '主観データ/主観その他/主観相関用.xlsx'
df = pd.read_excel(file_path, sheet_name='all')

# 列名の前後のスペースを削除
df.columns = df.columns.str.strip()

first_columns = ['物件名', '物件CD', 'テナントCD']

# エリアフラグを立てる
syuto_area = ['東京都', '神奈川県', '千葉県', '埼玉県']
df['首都圏フラグ'] = df['都道府県'].apply(lambda x: 1 if x in syuto_area else 0)

sanndaitoshi_area = ['東京都', '大阪府', '愛知県']
df['三大都市圏フラグ'] = df['都道府県'].apply(lambda x: 1 if x in sanndaitoshi_area else 0)

# 施設分類フラグ
df['施設分類フラグ'] = df['施設分類'].apply(lambda x: 1 if x == '事務所' else 0)

# 竣工日の年を取得
df['竣工年'] = pd.to_datetime(df['竣工日'], errors='coerce').dt.year

# 規模フラグ
df['規模フラグ'] = df['規模'].apply(lambda x: 1 if x == '大規模' else 0)

# 業種分類フラグの設定
industry_categories = [
    '農業、林業', '漁業', '鉱業，採石業，砂利採取業', '建設業', '製造業',
    '電気・ガス・熱供給・水道業', '情報通信業', '運輸業、郵便業', '卸売業、小売業',
    '金融業、保険業', '不動産業，物品賃貸業', '学術研究，専門・技術サービス業',
    '宿泊業，飲食サービス業', '生活関連サービス業，娯楽業', '教育，学習支援業',
    '医療、福祉', '複合サービス業務', 'サービス業（他に分類されない）', '公務（他に分類されない）'
]

for category in industry_categories:
    df[f'{category}フラグ'] = df['業種分類'].apply(lambda x: 1 if x == category else 0)

industry_columns = [f'{category}フラグ' for category in industry_categories]

# 勤務制度のフラグ
df['フレックスタイムフラグ'] = df['勤務制度'].apply(lambda x: 1 if x == 'フレックスタイム' else 0)
df['固定の労働時間制フラグ'] = df['勤務制度'].apply(lambda x: 1 if x == '固定の労働時間制' else 0)
df['その他フラグ'] = df['勤務制度'].apply(lambda x: 1 if x == 'その他' else 0)

# 着座率の計算
df['着座率'] = (df['合計在籍人数'] / df['座席数']) * 100

# 面積感覚フラグ
space_flags = {
    'かなり狭いと感じている': 'かなり狭いフラグ',
    'やや狭いと感じている': 'やや狭いフラグ',
    'ちょうど良いと感じている': 'ちょうど良いフラグ',
    'やや広いと感じている': 'やや広いフラグ',
    'かなり広いと感じている': 'かなり広いフラグ',
    'わからない': 'わからないフラグ'
}
for key, flag in space_flags.items():
    df[flag] = df['面積感覚'].apply(lambda x: 1 if x == key else 0)

# 入居中面積(㎡)の計算
if '入居中面積(坪)' in df.columns:
    df['入居中面積(㎡)'] = df['入居中面積(坪)'] * 3.30579

# サーベイ関連の合計点
survey_columns = [
    '防災', '防犯', '感染症対策', 'ダイバーシティ', '光', '温熱',
    '空気', '音', '清掃', '緑化', 'トイレ', 'エレベーター',
    '健康施策', '内装', 'ワークスペース', 'リフレッシュ',
    'コミュニケーション', '通信', '景観', '利便性', '地域愛着', 'リレーション'
]
df['合計点'] = df[survey_columns].sum(axis=1)

# スペース関連の列の正規化
renamed_space_columns = [
    '固定席', 'フリーアドレス席', 'グループアドレス席', 'オープンミーティングスペース',
    'リモート会議用ブース', '電話専用ブース', '集中スペース', '食堂カフェスペース',
    'リフレッシュスペース', 'コラボレーションスペース', 'その他スペース'
]
space_rename_dict = {
    'スペース_固定席': '固定席',
    'スペース_フリーアドレス席（個人が自由に選ぶことができるスタイルのデスク）': 'フリーアドレス席',
    'スペース_グループアドレス席（部署やチーム等の決められたエリアの中で、個人が自由に選ぶことができるスタイルのデスク）': 'グループアドレス席',
    'スペース_オープンなミーティングスペース': 'オープンミーティングスペース',
    'スペース_リモート会議用ブース・個室': 'リモート会議用ブース',
    'スペース_電話専用ブース・個室': '電話専用ブース',
    'スペース_集中するためのスペース': '集中スペース',
    'スペース_食堂・カフェスペース': '食堂カフェスペース',
    'スペース_リフレッシュスペース': 'リフレッシュスペース',
    'スペース_外部とのコラボレーションを目的としたスペース': 'コラボレーションスペース',
    'スペース_その他': 'その他スペース'
}
df.rename(columns=space_rename_dict, inplace=True)

# analysis_columnsの作成
analysis_columns = (
    first_columns + 
    ['竣工年', '首都圏フラグ', '三大都市圏フラグ', '施設分類フラグ', '規模フラグ', '延床（㎡）', 
    '年間受託金額(千円)', '月間受託金額（千円）'] + 
    renamed_space_columns + industry_columns +
    ['フレックスタイムフラグ', '固定の労働時間制フラグ', 'その他フラグ', '入居中面積(㎡)', '合計在籍人数', '出社率', '座席数', 
    'かなり狭いフラグ', 'やや狭いフラグ', 'ちょうど良いフラグ', 'やや広いフラグ', 'かなり広いフラグ', '着座率', '総合満足度'] +
    survey_columns + ['合計点'] 
)

# 選択した列がデータフレームに存在するかを確認
missing_columns = [col for col in analysis_columns if col not in df.columns]
if missing_columns:
    print(f"以下の列が見つかりませんでした: {missing_columns}")

# 出力ディレクトリの作成
output_dir = '主観データ'
os.makedirs(output_dir, exist_ok=True)

# データの出力
output_file_path = os.path.join(output_dir, '主観分析用.xlsx')
df[analysis_columns].to_excel(output_file_path, index=False)
print(f"データのクリーニングお疲れさまです。{output_file_path} に保存できました。")

