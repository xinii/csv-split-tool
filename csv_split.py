import pandas as pd

df = pd.read_excel('input.xlsx')

# ⭐ 先修正日期列
last_col = df.columns[-1]
df[last_col] = pd.to_datetime(
    df[last_col],
    origin='1899-12-30',
    unit='D',
    errors='coerce'
).dt.strftime('%Y年%m月%d日')

# ⭐ 再处理其它列
for col in df.columns[:-1]:
    df[col] = df[col].astype(str)

df = df.fillna('')
result = []

for _, row in df.iterrows():
    value = str(row.iloc[4]).strip().replace('　', '')
    value = value.zfill(8)

    if value == '02025160':
        row1 = row.copy()
        row1.iloc[4] = '01064773'

        row2 = row.copy()
        row2.iloc[4] = '02025110'

        result.extend([row1, row2])
    else:
        result.append(row)

new_df = pd.DataFrame(result)
# new_df.to_csv('output.csv', index=False, encoding="utf-8-sig", quoting=1)
new_df.to_excel('output.xlsx', index=False)
