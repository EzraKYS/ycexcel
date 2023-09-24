import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime

# 打开Excel文件
wb = openpyxl.load_workbook('ycexcel/异常20230918.xlsx')

# 选择要操作的工作表
ws = wb['Sheet1']

# 遍历每个单元格
for row in ws.iter_rows():
    for cell in row:
        # 居中显示整个表的内容
        cell.alignment = Alignment(horizontal='center', vertical='center')
        # 删除日期格式的内容，并删除本行
        if isinstance(cell.value, datetime):
            cell.value = ''
            
        # 替换'关于'和'列入经营异常名录的公告NEW'为空
        if cell.value and '关于' in cell.value:
            cell.value = cell.value.replace("关于", "").replace("列入经营异常名录的公告\xa0NEW", "")
        # elif cell.value and '市场监督管理局' in cell.value:
        #     # 删除包含'市场监督管理局'的内容，并删除本行
        #     cell.value = ''
        else:
            # 删除其它内容
            cell.value = ''

for row in reversed(list(ws.rows)):
    if all(cell.value is None or cell.value == '' for cell in row):        
        ws.delete_rows(row[0].row)
        

# 保存Excel文件
wb.save('yc20230918.xlsx')