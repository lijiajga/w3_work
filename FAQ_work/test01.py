import openpyxl

# 加载Excel文件和特定Sheet
file_path = 'FAQ_v2.xlsx'
sheet_name = 'Sheet'
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook[sheet_name]

# 遍历行，检查K列的值，并根据条件在A列中追加后缀
for row in worksheet.iter_rows(min_row=2):  # 假设第一行是标题
    k_value = row[10].value  # K列索引为10（基于0的索引）
    a_cell = row[0]  # A列索引为0

    # 根据K列的值来判断并处理A列的内容
    if k_value == 'EXCEL':
        a_cell.value = f'{a_cell.value}.xlsx'
    elif k_value == 'PPT':
        a_cell.value = f'{a_cell.value}.pptx'

# 保存修改后的Excel文件
workbook.save('FAQ_v2_updated.xlsx')
