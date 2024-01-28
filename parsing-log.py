# -*- coding: utf-8 -*-
"""
Created on Mon Dec 18 11:26:11 2023

@author: chienhuak
"""

import os
import pandas as pd
import zipfile

def list_files_and_directories_to_csv(path, output_csv_path):
    try:
        # 使用os.walk獲取目錄結構
        data = {'目錄': [], '檔案名稱': [], '副檔名': [], '反斜線數量': []}
        for root, dirs, files in os.walk(path):
            for file in files:
                data['目錄'].append(root)
                data['檔案名稱'].append(os.path.splitext(file)[0])  # 取得檔案名稱
                data['副檔名'].append(os.path.splitext(file)[1])  # 取得副檔名
                data['反斜線數量'].append(root.count('\\'))  # 計算反斜線數量

        # 檢查是否為 zip 檔案，如果是，則加入 zip 內的檔案
        if path.endswith('.zip'):
            with zipfile.ZipFile(path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    file_path = file_info.filename
                    data['目錄'].append(os.path.dirname(file_path))
                    data['檔案名稱'].append(os.path.splitext(os.path.basename(file_path))[0])
                    data['副檔名'].append(os.path.splitext(os.path.basename(file_path))[1])
                    data['反斜線數量'].append(file_path.count('/'))  # 計算反斜線數量（zip 內使用 /）

        # 將數據轉換成 DataFrame
        df = pd.DataFrame(data)

        # 將 DataFrame 寫入 CSV 檔案
        df.to_csv(output_csv_path, index=False)

        print(f'已將檔案清單保存到 {output_csv_path}')

    except Exception as e:
        print(f'發生錯誤: {e}')

# 輸入你的路徑
path = r'\\cgcfsii\Confidential\97_Test Result\Google\NBPB00\TTT'  # 請替換為你的實際路徑

# 指定要保存的 CSV 檔案路徑
output_csv_path = r'C:\Users\chienhuak\Desktop\testK_zip.csv'  # 請替換為你想要保存的路徑

# 呼叫函數，將檔案清單保存到 CSV 檔案
list_files_and_directories_to_csv(path, output_csv_path)
