# -*- coding: utf-8 -*-
"""
Created on Thu Dec 14 11:29:32 2023

@author: chienhuak
"""

import os
import pandas as pd
import re

def extract_folder_name(path):
    # 使用正則表達式擷取路徑中 "TT\"到下一個"\"中間的文字
    match = re.search(r'\\TT\\([^\\]+)', path)
    if match:
        return match.group(1)
    else:
        return ''  # 如果找不到匹配，返回空字符串

def count_backslashes(path):
    # 計算路徑中反斜線的數量
    return path.count('\\')

def list_files_and_directories_to_csv(path, output_csv_path):
    try:
        # 使用os.walk獲取目錄結構
        data = {'目錄': [], '檔案名稱': [], '副檔名': [], '擷取的資訊': [], '反斜線數量': []}
        for root, dirs, files in os.walk(path):
            for file in files:
                data['目錄'].append(root)
                data['檔案名稱'].append(os.path.splitext(file)[0])  # 取得檔案名稱
                data['副檔名'].append(os.path.splitext(file)[1])  # 取得副檔名
                data['擷取的資訊'].append(extract_folder_name(root))  # 擷取資訊
                data['反斜線數量'].append(count_backslashes(root))  # 計算反斜線數量

        # 將數據轉換成DataFrame
        df = pd.DataFrame(data)

        # 將 DataFrame 寫入 CSV 檔案
        df.to_csv(output_csv_path, index=False)

        print(f'已將檔案清單保存到 {output_csv_path}')

    except Exception as e:
        print(f'發生錯誤: {e}')

# 輸入你的路徑
path = r'\\cgcfsii\Confidential\97_Test Result\Google\NBP90X\TTT'  # 請替換為你的實際路徑

# 指定要保存的 CSV 檔案路徑
output_csv_path = r'C:\Users\chienhuak\Desktop\testG.csv'  # 請替換為你想要保存的路徑

# 呼叫函數，將檔案清單保存到 CSV 檔案
list_files_and_directories_to_csv(path, output_csv_path)
