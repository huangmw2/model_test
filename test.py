from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
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

def find_merged_cell(ws, target_value):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == target_value:
                # 遍历所有合并单元格，检查该单元格是否在合并区域内
                for merged_range in ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        start_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col).coordinate
                        merged_columns = merged_range.max_col - merged_range.min_col + 1
                        merged_rows = merged_range.max_row - merged_range.min_row + 1
                        return {"start_cell": start_cell, "merged_columns": merged_columns, "merged_rows": merged_rows}
                # 如果该单元格没有合并，则返回它自身
                return {"start_cell": cell.coordinate, "merged_columns": 1, "merged_rows": 1}
    return None  # 未找到该值

# 测试
value = "确认结构和外观测试"
result = find_merged_cell(ws, value)

if result:
    print(f"起始单元格: {result['start_cell']}, 合并列数: {result['merged_columns']}, 合并行数: {result['merged_rows']}")
else:
    print("未找到该值")

    def get_topic_range(self, topic_name):
    for row in self.ws.iter_rows():
        for cell in row:
            if cell.value == topic_name:
                # 遍历所有合并单元格，检查该单元格是否在合并区域内
                for merged_range in self.ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        start_cell = self.ws.cell(row=merged_range.min_row, column=merged_range.min_col).coordinate
                        range_str = f"{self.ws.cell(row=merged_range.min_row, column=merged_range.min_col).coordinate}:" \
                                    f"{self.ws.cell(row=merged_range.max_row, column=merged_range.max_col).coordinate}"
                        
                        # 计算右侧一列的新单元格坐标
                        new_col_index = merged_range.max_col + 1  # 右侧一列
                        next_column_cell = self.ws.cell(row=merged_range.min_row, column=new_col_index).coordinate

                        return next_column_cell, range_str  # 返回右侧单元格和合并范围

                # 如果该单元格没有合并，则计算右侧一列的新单元格
                next_column_cell = self.ws.cell(row=cell.row, column=cell.column + 1).coordinate
                return next_column_cell, f"{cell.coordinate}:{cell.coordinate}"  # 单个单元格范围就是它自己

    return None, None  # 未找到该值
