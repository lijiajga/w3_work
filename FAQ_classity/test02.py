
#只根据文件名称进行匹配，根据V方式匹配

FAQ_checklist_message =[]
FAQ_KCP_message =[]
FAQ_SOD_message =[]
FAQ_模板_message =[]
FAQ_行政发文_message =[]

import pandas as pd
import os
import shutil
import openpyxl
from test_V_name import find_V , find_V_row

#复制文件
def copy_file(input_path,filename,output_path):
    shutil.copy(os.path.join(input_path, filename), output_path)


#判断w3文件的类型
def get_all_list():
    file_path = r'D:\实体文件_汇总\w3文件收集规则\ISC+知识分析0716_v2.xlsx'
    sheet_name = '线上线下收集模板'
    data = pd.read_excel(file_path, sheet_name=sheet_name)
    data = data[data['归档位置'] == 'W3']
    for index, row in data.iterrows():
        if row['小类'] == 'Checklist':
            FAQ_checklist_message.append(find_V_row(row['文档名称（当前）']))
        elif row['小类'] == 'KCP':
            FAQ_KCP_message.append(find_V_row(row['文档名称（当前）']))
        elif row['小类'] == 'SOD':
            FAQ_SOD_message.append(find_V_row(row['文档名称（当前）']))
        elif row['小类'] == '模板':
            FAQ_模板_message.append(find_V_row(row['文档名称（当前）']))
        elif row['小类'] == '行政发文':
            FAQ_行政发文_message.append(find_V_row(row['文档名称（当前）']))
#文件归档到指定的文件夹中
def file_classify():
    # 定义文件夹路径
    folder_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3文件_总'
    # 获取文件夹中的所有文件名
    w3_file_names = os.listdir(folder_path)

    for file in w3_file_names:
        if find_V(file) in FAQ_checklist_message:
            copy_file(folder_path,file,r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_知识库(强匹配)\FAQ细分\FAQ_checklist')
        
        if find_V(file) in FAQ_KCP_message:
            copy_file(folder_path,file,r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_知识库(强匹配)\FAQ细分\FAQ_KCP')
        
        if find_V(file) in FAQ_SOD_message:
            copy_file(folder_path,file,r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_知识库(强匹配)\FAQ细分\FAQ_SOD')
        
        if find_V(file) in FAQ_模板_message:
            copy_file(folder_path,file,r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_知识库(强匹配)\FAQ细分\FAQ_模板')
        
        if find_V(file) in FAQ_行政发文_message:
            copy_file(folder_path,file,r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_知识库(强匹配)\FAQ细分\FAQ_行政发文')
            
    print("分类保存内容已完成！")


get_all_list()
file_classify()
