# 対象となる文字列
target = '0123456789'

# findで検索
i = target.find('56')
if i >= 0:
    print('(*1) 0から数えて', i, '文字目にあります')
else:
    print('(*1) 見つかりません')

# inで検索
if '56' in target:
    print('(*2) 見つかりました')
else:
    print('(*2) 見つかりません')
