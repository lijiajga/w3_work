import openpyxl

# 加载Excel文件和特定Sheet
file_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\ISC+知识分析0716_v5.xlsx'
sheet_name = '线上线下收集模板'
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook[sheet_name]

# 使用set来存储已记录的C列值
seen_c_values = set()
rows_to_delete = []

# 遍历行，记录需要删除的行
for row in worksheet.iter_rows(min_row=2, values_only=False):  # 假设第一行为标题
    e_value = row[4].value  # E列索引为4（基于0的索引）
    c_value = row[2].value  # C列索引为2

    if e_value == 'W3':
        # 如果C列值已经存在，则标记该行供删除
        if c_value in seen_c_values:
            rows_to_delete.append(row[0].row)
        else:
            seen_c_values.add(c_value)

# 删除标记的行，从最后一行开始删除以避免索引变化
for row_index in sorted(rows_to_delete, reverse=True):
    worksheet.delete_rows(row_index)

# 保存修改后的Excel文件
workbook.save('ISC+知识分析0716_v5_update.xlsx')
