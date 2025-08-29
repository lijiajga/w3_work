import os
import re

# 定义要处理的文件夹路径
folder_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_zip\w3_all_test'

# 获取文件夹中的所有文件
files = os.listdir(folder_path)

# 用于存储文件的最新版本信息
latest_files = {}

# 正则表达式匹配带版本号的文件名，支持多种文件类型
pattern = re.compile(r'^(.*?)(V\d+\.\d+).*?(\.docx|\.doc|\.xlsx|\.pptx)$', re.IGNORECASE)

# 遍历所有文件
for file in files:
    match = pattern.match(file)
    if match:
        name_prefix = match.group(1)
        version = match.group(2)
        if name_prefix not in latest_files:
            latest_files[name_prefix] = (file, version)
        else:
            # 比较版本号
            existing_version = latest_files[name_prefix][1]
            if version > existing_version:
                # 更新最新版本信息
                latest_files[name_prefix] = (file, version)

# 删除非最新版本文件
for file in files:
    match = pattern.match(file)
    if match:
        name_prefix = match.group(1)
        if name_prefix in latest_files and file != latest_files[name_prefix][0]:
            os.remove(os.path.join(folder_path, file))
            print(f'删除文件：{file}')

print("任务完成")
