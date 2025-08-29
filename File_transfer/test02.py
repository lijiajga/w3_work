import os
import shutil

# 定义文件夹路径
file1_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3知识库实际上传\ISC+通用-业务操作指导书-内部公开'
file2_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_知识库(强匹配)\ISC+通用-业务操作指导书-内部公开'
file3_path = r'D:\Pan_xfusion_worksapce\1.《《任务》》\7月\1.知识库上传\7.w3\w3_知识库(强匹配)\比《现有知识库》匹配多出来的文件\ISC+通用-业务操作指导书-内部公开'

def copy_nonexistent_files(file1_path, file2_path, file3_path):
    copied_file_count = 0
    # 获取file1和file2文件夹中的所有文件
    files_in_file1 = set(os.listdir(file1_path))
    files_in_file2 = os.listdir(file2_path)
    
    # 判断并复制不存在于file1中的文件到file3
    for file in files_in_file2:
        if file not in files_in_file1:
            source_file_path = os.path.join(file2_path, file)
            destination_file_path = os.path.join(file3_path, file)
            shutil.copy(source_file_path, destination_file_path)
            copied_file_count += 1
            # print(f'复制文件 {file} 到 {file3_path}')
    print(copied_file_count)


# 执行复制操作
copy_nonexistent_files(file1_path, file2_path, file3_path)
