import openpyxl

# 示例数组
a = [1, 2, 3, 4, 5]

# 创建一个新的 Excel 工作簿和工作表
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "sheet1"

# 循环写入数组 a 的元素，从单元格 A2 开始
for i, value in enumerate(a, start=2):
    ws[f'A{i}'] = value

# 保存 Excel 文件
excel_file_path = 'G:\project\LHSextract\三会一课未上传/array_to_excel.xlsx'
wb.save(excel_file_path)

print(f'Excel 文件已保存到 {excel_file_path}')
