# ダミーデータ
dummy_data = [
    ["伊藤",300],
    ["伊藤",600],
    ["伊藤",200],
    ["田中",300],
    ["田中",200]
]
# データを顧客名で辞典型に分割
users = {}
for row in dummy_data:
    name,value = row
    # 顧客名が初出だった場合　リストを初期化
    if name not in users:
        users[name] = []
    # 顧客名をキーにして、値にデータを追記
    users[name].append(row)
    
print(users) # {'伊藤': [['伊藤', 300], ['伊藤', 600], ['伊藤', 200]], '田中': [['田中', 300], ['田中', 200]]}

for name,rows in users.items(): #　顧客ごとに集計 
    print(users.items(),"items")# ([('伊藤', [['伊藤', 300], ['伊藤', 600], ['伊藤', 200]]), ('田中', [['田中', 300], ['田中', 200]])])
    #　顧客の購入金額を合計
    total = 0
    for row in rows:
        total += row[1]
    #print(name,total)