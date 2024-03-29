# -*- coding: utf-8 -*-
"""
Created on Fri Dec 22 11:56:30 2023

@author: chienhuak
"""


import os
import pandas as pd
from bs4 import BeautifulSoup

def count_backslashes(path):
    # 計算路徑中反斜線的數量
    return path.count('\\')

def extract_folder_name(path):
    # 尋找 "TT\" 的位置
    start_index = path.find('TT\\')
    
    # 如果找到 "TT\"，則在這個位置後尋找下一個反斜線 "\"
    if start_index != -1:
        end_index = path.find('\\', start_index + 3)
        if end_index != -1:
            return path[start_index + 3:end_index]
    
    return ''  # 如果找不到匹配，返回空字符串

def extract_info_from_K_250C_html(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            soup = BeautifulSoup(content, 'html.parser')

            # 尋找 Test case ID
            test_case_id_tag = soup.select_one('a#Title1')
            test_case_id = test_case_id_tag.text.strip() if test_case_id_tag else ''

            # 尋找 Band
            band_tag = soup.select_one('th:contains("Operating Band") + td')
            band = band_tag.text.strip() if band_tag else ''


            # 尋找 Configuration
            voltage = soup.select_one('th:contains("Voltage") + td').text.strip()
            temperature = soup.select_one('th:contains("Temperature") + td').text.strip()
            configuration = f"{voltage} - {temperature}"

            # 尋找 Result
            result_tag = soup.select_one('#TestCaseVerdict')
            result = result_tag.text.strip() if result_tag else ''


            return test_case_id, band, configuration, result

    except Exception as e:
        print(f'解析 HTML 時發生錯誤: {e}')
        return '', '', '', ''

def extract_info_from_K195html(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            soup = BeautifulSoup(content, 'html.parser')

            # 尋找 Test case ID
            test_case_id_tag = soup.find('th', {'scope': 'col'}, text='Test Case Number and Name:')
            test_case_id = test_case_id_tag.find_next('td').text.split('|')[-1].strip() if test_case_id_tag else ''
#            test_case_id = test_case_id_tag.find_next('td').text.strip() if test_case_id_tag else ''


            # 尋找 Band
            band_tag = soup.select_one('#divTestSummary table.summary-table tbody tr.summary-tr-body.border-bottom td:nth-child(6)')
            band = band_tag.text.strip() if band_tag else ''


            # 尋找 Configuration
            configuration_tag = soup.select_one('#divTestSummary table.summary-table tbody tr.summary-tr-body.border-bottom td:nth-child(5)')
            configuration = configuration_tag.text.strip() if band_tag else ''

            # 尋找所有符合條件的 tr 元素
            # rows = soup.find_all('tr', class_='summary-tr-body border-bottom')


#            configuration_tag = soup.find('th', {'scope': 'col'}, text='Test Case Title:')
#            configuration = configuration_tag.find_next('td').text.strip() if configuration_tag else ''

            # 尋找 Result
            result_tag = soup.find('th', {'scope': 'col'}, text='Test Case Verdict:')
            result = result_tag.find_next('td').text.strip() if result_tag else ''

            return test_case_id, band, configuration, result

    except Exception as e:
        print(f'解析 HTML 時發生錯誤: {e}')
        return '', '', '', ''


def extract_info_from_168html(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            soup = BeautifulSoup(content, 'html.parser')

            # 尋找 Test case ID
            test_case_id_tag = soup.find('th', {'scope': 'col'}, text='Test Case Number and Name:')
            test_case_id = test_case_id_tag.find_next('td').text.split('|')[-1].strip() if test_case_id_tag else ''
#            test_case_id = test_case_id_tag.find_next('td').text.strip() if test_case_id_tag else ''


            # 尋找 Band
            band_tag = soup.select_one('#divTestSummary table.summary-table tbody tr.summary-tr-body.border-bottom td:nth-child(6)')
            band = band_tag.text.strip() if band_tag else ''


            # 尋找 Configuration
            configuration_tag = soup.select_one('#divTestSummary table.summary-table tbody tr.summary-tr-body.border-bottom td:nth-child(5)')
            configuration = configuration_tag.text.strip() if band_tag else ''

            # 尋找所有符合條件的 tr 元素
            # rows = soup.find_all('tr', class_='summary-tr-body border-bottom')


#            configuration_tag = soup.find('th', {'scope': 'col'}, text='Test Case Title:')
#            configuration = configuration_tag.find_next('td').text.strip() if configuration_tag else ''

            # 尋找 Result
            result_tag = soup.find('th', {'scope': 'col'}, text='Test Case Verdict:')
            result = result_tag.find_next('td').text.strip() if result_tag else ''

            return test_case_id, band, configuration, result

    except Exception as e:
        print(f'解析 HTML 時發生錯誤: {e}')
        return '', '', '', ''

def extract_info_from_html(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            soup = BeautifulSoup(content, 'html.parser')

            # 尋找 Test case ID
            test_case_id_tag = soup.find('div', id='title')
            test_case_id = test_case_id_tag.text.split('—')[0].strip() if test_case_id_tag and test_case_id_tag.text else ''

            # 尋找 Band 和 Configuration
            config_tag = soup.find('div', id='summary_configuration')
            config_text = config_tag.text if config_tag and config_tag.text else ''
            band, configuration = map(str.strip, config_text.split(';')[:2])


            # 尋找 Result
            result_tag = soup.find('div', id='summary_verdict')
            result_text = result_tag.text.strip() if result_tag and result_tag.text else ''
            
            
            return test_case_id, band, configuration, result_text

    except Exception as e:
        print(f'解析 HTML 時發生錯誤: {e}')
        return '', '', '', ''


def list_files_and_directories_to_csv(path, output_csv_path):
    try:
        # 使用os.walk獲取目錄結構
        data = {'Path': [], 'File Name': [], 'File Types': [], 'TP': [],'Test case ID': [], 'Band': [], 'Configuration': [], 'Result': [], 'Count /': []}
        for root, dirs, files in os.walk(path):
            # 過濾250(c)檔案名稱包含 'TS' 的檔案
            TS_files = [file for file in files if 'TS ' in file]
            for file in TS_files:
                file_path = os.path.join(root, file)
                data['Path'].append(root)
                data['File Name'].append(os.path.splitext(file)[0])  # 取得檔案名稱
                data['File Types'].append(os.path.splitext(file)[1])  # 取得副檔名
                data['TP'].append(extract_folder_name(root))  # 擷取資訊

                # 擷取 HTML 中的資訊
                test_case_id, band, configuration, result = extract_info_from_K_250C_html(file_path)

                data['Test case ID'].append(test_case_id)
                data['Band'].append(band)
                data['Configuration'].append(configuration)
                data['Result'].append(result)
                data['Count /'].append(count_backslashes(root))  # 計算反斜線數量
            
            # 過濾檔案名稱包含 '3GPP' 的檔案
            gpp_report_files = [file for file in files if '3GPP' in file]
            for file in gpp_report_files:
                file_path = os.path.join(root, file)
                data['Path'].append(root)
                data['File Name'].append(os.path.splitext(file)[0])  # 取得檔案名稱
                data['File Types'].append(os.path.splitext(file)[1])  # 取得副檔名
                data['TP'].append(extract_folder_name(root))  # 擷取資訊

                # 擷取 HTML 中的資訊
                test_case_id, band, configuration, result = extract_info_from_K195html(file_path)

                data['Test case ID'].append(test_case_id)
                data['Band'].append(band)
                data['Configuration'].append(configuration)
                data['Result'].append(result)
                data['Count /'].append(count_backslashes(root))  # 計算反斜線數量


            # 過濾檔案名稱包含 'TCReport' 的檔案
            tc_report_files = [file for file in files if 'TCReport' in file]
            for file in tc_report_files:
                file_path = os.path.join(root, file)
                data['Path'].append(root)
                data['File Name'].append(os.path.splitext(file)[0])  # 取得檔案名稱
                data['File Types'].append(os.path.splitext(file)[1])  # 取得副檔名
                data['TP'].append(extract_folder_name(root))  # 擷取資訊

                # 擷取 HTML 中的資訊
                test_case_id, band, configuration, result = extract_info_from_168html(file_path)

                data['Test case ID'].append(test_case_id)
                data['Band'].append(band)
                data['Configuration'].append(configuration)
                data['Result'].append(result)
                data['Count /'].append(count_backslashes(root))  # 計算反斜線數量

            # 過濾檔案名稱為'OnlineReport'的檔案
            online_report_files = [file for file in files if file.lower() == 'onlinereport.htm']
            for file in online_report_files:
                file_path = os.path.join(root, file)
                data['Path'].append(root)
                data['File Name'].append(os.path.splitext(file)[0])  # 取得檔案名稱
                data['File Types'].append(os.path.splitext(file)[1])  # 取得副檔名
                data['TP'].append(extract_folder_name(root))  # 擷取資訊

                # 擷取 HTML 中的資訊
                test_case_id, band, configuration, result = extract_info_from_html(file_path)

                data['Test case ID'].append(test_case_id)
                data['Band'].append(band)
                data['Configuration'].append(configuration)
                data['Result'].append(result)
                data['Count /'].append(count_backslashes(root))  # 計算反斜線數量

        # 將數據轉換成 DataFrame
        df = pd.DataFrame(data)

        # 將 DataFrame 寫入 CSV 檔案
        df.to_csv(output_csv_path, index=False)

        print(f'已將檔案清單保存到 {output_csv_path}')

    except Exception as e:
        print(f'發生錯誤: {e}')


# 輸入你的路徑
path = r'C:\Users\chienhuak\Downloads\Log download\CM4\NBPF04\Request_240115'  # 請替換為你的實際路徑

# 按需指定要保存的 CSV 檔案路徑
output_csv_path = r'C:\Users\chienhuak\Desktop\testO_NBPF04_240117.csv'  # 請替換為你想要保存的路徑

# 呼叫函數，將檔案清單保存到 CSV 檔案
list_files_and_directories_to_csv(path, output_csv_path)
