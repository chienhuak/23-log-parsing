# -*- coding: utf-8 -*-
"""
Created on Wed Dec 13 15:26:22 2023

@author: chienhuak
"""

import os
import pandas as pd

def list_files_and_directories_to_excel(path, output_excel_path):
    try:
        # 使用os.walk獲取目錄結構
        data = {'目錄': [], '檔案': []}
        for root, dirs, files in os.walk(path):
            for file in files:
                data['目錄'].append(root)
                data['檔案'].append(file)

        # 將數據轉換成DataFrame
        df = pd.DataFrame(data)

        # 拆分 DataFrame 為多個小的 DataFrame
        chunk_size = 1048575  # Excel 表格的行數上限
        chunks = [df[i:i + chunk_size] for i in range(0, df.shape[0], chunk_size)]

        # 建立 Excel writer
        with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
            for idx, chunk in enumerate(chunks):
                # 將每個小 DataFrame 寫入 Excel 工作表
                chunk.to_excel(writer, sheet_name=f'Sheet{idx + 1}', index=False)

        print(f'已將檔案清單保存到 {output_excel_path}')

        # 打開 Excel 檔案
        os.system(output_excel_path)

    except Exception as e:
        print(f'發生錯誤: {e}')

# 輸入你的路徑
path = r'\\cgcfsii\Confidential\97_Test Result\Google'  # 請替換為你的實際路徑

# 指定要保存的 Excel 檔案路徑
output_excel_path = r'C:\Users\chienhuak\Desktop\testH.xlsx'  # 請替換為你想要保存的路徑

# 呼叫函數，將檔案清單保存到 Excel 檔案並打開
list_files_and_directories_to_excel(path, output_excel_path)
