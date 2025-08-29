import os
import zipfile

# 定义压缩包所在的文件夹路径
zip_folder_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_zip'

# 定义解压后的目标文件夹路径
unzip_folder_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_zip\w3_all_test'

# 如果目标文件夹不存在则创建
if not os.path.exists(unzip_folder_path):
    os.makedirs(unzip_folder_path)

# 获取压缩包列表
zip_files = [f for f in os.listdir(zip_folder_path) if f.endswith('.zip')]
# print(zip_files)
# 解压每个压缩包
for zip_file in zip_files:
    zip_file_path = os.path.join(zip_folder_path, zip_file)

    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        for file_info in zip_ref.infolist():
            file_info.filename = file_info.filename.encode('cp437').decode('gbk')
            # 解压到目标文件夹
            zip_ref.extract(file_info,unzip_folder_path)
    print(f'Extracted: {zip_file} to {unzip_folder_path}')
