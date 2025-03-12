from openpyxl import load_workbook
import sys
import os
# 获取运行目录
if getattr(sys, 'frozen', False):  # 检测是否是 PyInstaller 打包后的环境
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

#file_path = r"Data\.."
file_dir = os.path.join(base_path, "Data")
# 加载 Excel 文件
file_path = file_dir + r"\原文件.xlsx"
wb = load_workbook(file_path)
ws = wb.active  # 获取默认工作表

# 遍历表格查找 "测试项目" 的列索引
col_idx = None
row_idx = None

for row in ws.iter_rows():
    for cell in row:
        if cell.value == "测试项目":
            col_idx = cell.column  # 获取列索引（数字）
            row_idx = cell.row     # 获取行索引
            break
    if col_idx:
        break

# 获取该列 "测试项目" 下面的所有数据
column_data = [ws.cell(row=i, column=col_idx).value for i in range(row_idx + 1, ws.max_row + 1) if ws.cell(row=i, column=col_idx).value is not None]

print(column_data)
