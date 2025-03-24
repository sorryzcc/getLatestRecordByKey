import pandas as pd

# 读取Excel文件
file_path = 'mainland_textApplicationForm.xlsx'
df = pd.read_excel(file_path)

# 确保'inGameKey'和'mtime'列名正确无误，如果在Excel中不是这些名字，请根据实际情况修改
# 假设G列为'inGameKey'，H列为'mtime'
df.rename(columns={'客户端读取的key(InGameKey)': 'inGameKey', '修改时间(_mtime)': 'mtime'}, inplace=True)

# 将'mtime'转换为datetime类型以便比较
df['mtime'] = pd.to_datetime(df['mtime'])

# 按'inGameKey'分组，并应用agg函数来保留每个'inGameKey'最新的记录（即'mtime'最大的）
latest_entries = df.loc[df.groupby('inGameKey')['mtime'].idxmax()]

# 输出结果查看
print(latest_entries)

# 如有需要，将筛选后的数据保存到新的Excel文件
output_file_path = 'filtered_mainland_textApplicationForm.xlsx'
latest_entries.to_excel(output_file_path, index=False)