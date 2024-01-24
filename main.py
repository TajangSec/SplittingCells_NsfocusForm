import pandas as pd
import sys

if len(sys.argv) > 1:
    file_path = sys.argv[1]
    new_path=file_path[:-5]+"-IP展开.xlsx"
else:
    print("No file dragged.")
# 读取Excel
df = pd.read_excel(file_path, sheet_name='漏洞信息', index_col=1)
# 删除第一行、第一列
df = df.drop(labels='漏洞分布', axis=1)
df = df.drop(labels='序号')
# 将Unname 1等列名重命名
df.columns = ['危险程度', '漏洞名称', '服务分类', '应用分类', '系统分类', '威胁分类', '时间分类', 'CVE年份分类',
              '影响IP', '出现次数', 'CVE ID']
# 拆分单元格
df = df.drop(labels='影响IP', axis=1).join(
    df['影响IP'].str.split(' ', expand=True).stack().reset_index(level=1, drop=True).rename('影响IP'))
# 移动新列的位置
IPdata = df.pop('影响IP')
df.insert(8, '影响IP', IPdata)
df.reset_index()
# 写入新的Excel
writer = pd.ExcelWriter(new_path)
df.to_excel(writer, 'Sheet1')
writer.close()
