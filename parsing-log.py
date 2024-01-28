# -*- coding: utf-8 -*-
"""
Created on Mon Dec 18 15:33:46 2023

@author: chienhuak
"""

import os
import pandas as pd
from bs4 import BeautifulSoup

def extract_folder_name(path):
    # 尋找 "TT\" 的位置
    start_index = path.find('TT\\')
    
    # 如果找到 "TT\"，則在這個位置後尋找下一個反斜線 "\"
    if start_index != -1:
        end_index = path.find('\\', start_index + 3)
        if end_index != -1:
            return path[start_index + 3:end_index]
    
    return ''  # 如果找不到匹配，返回空字符串

def count_backslashes(path):
    # 計算路徑中反斜線的數量
    return path.count('\\')

def extract_info_from_html(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            soup = BeautifulSoup(content, 'html.parser')

            # 尋找 Test case ID
            test_case_id_tag = soup.find('div', id='title')
            test_case_id = test_case_id_tag.text.split('—')[-1].strip() if test_case_id_tag else ''

            # 尋找 Band 和 Configuration
            config_tag = soup.find('div', id='summary_configuration')
            config_text = config_tag.text if config_tag else ''
            band, configuration = map(str.strip, config_text.split(';')[:2])

            # 尋找 Result
            result_tag = soup.find('p', class_='result')
            result = result_tag.find('span', class_='verdict_fail').text.strip() if result_tag else ''

            return test_case_id, band, configuration, result

    except Exception as e:
        print(f'解析 HTML 時發生錯誤: {e}')
        return '', '', '', ''

def list_files_and_directories_to_csv(path, output_csv_path):
    try:
        # 使用os.walk獲取目錄結構
        data = {'目錄': [], '檔案名稱': [], '副檔名': [], 'Test case ID': [], 'Band': [], 'Configuration': [], 'Result': [], '反斜線數量': []}
        for root, dirs, files in os.walk(path):
            for file in files:
                file_path = os.path.join(root, file)
                data['目錄'].append(root)
                data['檔案名稱'].append(os.path.splitext(file)[0])  # 取得檔案名稱
                data['副檔名'].append(os.path.splitext(file)[1])  # 取得副檔名
                data['反斜線數量'].append(count_backslashes(root))  # 計算反斜線數量

                # 擷取 HTML 中的資訊
                test_case_id, band, configuration, result = extract_info_from_html(file_path)

                data['Test case ID'].append(test_case_id)
                data['Band'].append(band)
                data['Configuration'].append(configuration)
                data['Result'].append(result)
                

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
output_csv_path = r'C:\Users\chienhuak\Desktop\testL.csv'  # 請替換為你想要保存的路徑

# 呼叫函數，將檔案清單保存到 CSV 檔案
list_files_and_directories_to_csv(path, output_csv_path)
