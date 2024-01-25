# -*- coding: utf-8 -*-
"""
Created on Sat Dec  9 16:03:10 2023

@author: chienhuak
"""

import os
import pandas as pd

def list_files_and_directories_to_excel(path, output_excel_path):
    try:
        data = {'目錄': [], '檔案': []}

        # 使用os.walk獲取目錄結構
        for root, dirs, files in os.walk(path):
            # 列印目錄
            for file in files:
                data['目錄'].append(root)
                data['檔案'].append(file)

        # 將數據轉換成DataFrame
        df = pd.DataFrame(data)

        # 將DataFrame保存到Excel檔案
        df.to_excel(output_excel_path, index=False)
        print(f'已將檔案清單保存到 {output_excel_path}')

        # 打開Excel檔案
        os.system(output_excel_path)

    except Exception as e:
        print(f'發生錯誤: {e}')

# 輸入你的路徑
path = r'\\cgcfsii\Confidential\97_Test Result\Google\NBP90X\TTT'  # 請替換為你的實際路徑

# 指定要保存的Excel檔案路徑
output_excel_path = r'C:\Users\chienhuak\Desktop\testG.xlsx'  # 請替換為你想要保存的路徑

# 呼叫函數，將檔案清單保存到Excel檔案並打開
list_files_and_directories_to_excel(path, output_excel_path)
