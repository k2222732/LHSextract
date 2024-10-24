import os
import pandas as pd

# 指定文件夹路径
folder_path = 'C:\\project\\LHSextract'
merged_data = []

def merge():
    # 遍历文件夹中的所有文件
    for filename in os.listdir(folder_path):
        if "班子成员信息" in filename and filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            # 读取Excel文件，保留表头
            df = pd.read_excel(file_path, dtype={'公民身份证号码': str})
            merged_data.append(df)

    # 合并所有数据
    if merged_data:
        result = pd.concat(merged_data, ignore_index=True)
        # 保存合并后的数据
        result.to_excel(os.path.join(folder_path, '支部班子成员信息.xlsx'), index=False)