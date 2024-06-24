# %%
import os
import pandas as pd
from time import sleep
from multiprocessing import Pool, cpu_count
from glob import glob

from tqdm import tqdm
import numpy as np
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def file_process(path):
    idx = pd.read_excel(path)
    try:
        idx = idx[idx['Unnamed: 2'].str.replace(' ', '') == '公司名稱'].index[0]
        return pd.read_excel(path, header=idx+1)
    except:
        print(path)


def data_process(dataframe):
    try:
        dataframe.columns = dataframe.columns.str.replace('\n', '')
        dataframe.columns = dataframe.columns.str.replace('　　', '')
        dataframe.columns = dataframe.columns.str.replace(' ', '')
    except:
        return
    if 'Unnamed:13' in dataframe.columns:
        dataframe = dataframe.drop(columns=['Unnamed:13'])
    
    if dataframe['公司型態'].dtype == 'int64':
        dataframe['公司型態'] = dataframe['公司型態'].map({1: '上市', 
                                                        2: '上櫃', 
                                                        3: '未上市(櫃)',
                                                        4: '補辦公開發行',
                                                        5: '興櫃',
                                                        6: '外國企業'})
    dataframe[['承銷商', '發行價格']] = dataframe[['承銷商', '發行價格']].replace(' ', np.nan)
    return dataframe.rename(columns={'金額(元)': '金額',
                                     '申報生效日期': '生效日期',
                                     '備註1': '案件性質',
                                     '申報事項': '案件類別'})


# %%
if __name__ == '__main__':
    if not os.path.exists('tmp'):
        os.makedirs('tmp')

    options = Options()
    appState = {
        'recentDestinations': [{
            'id': 'Save as PDF',
            'origin': 'local',
            'account': ''
        }],
        'selectedDestinationId': 'Save as PDF',
        'version': 2
    }
    prefs = {
        "download.default_directory": os.path.join(os.getcwd(), 'tmp'),
    }
    options.add_experimental_option('prefs', prefs)
    options.add_argument('--kiosk-printing')

    driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
    driver.get(f"https://www.sfb.gov.tw/ch/home.jsp?id=95&parentpath=0,2")
    driver.execute_script("advqry();")
    driver.execute_script("""document.getElementById("keyword").value="每日";""")
    driver.execute_script("qry();")
    sleep(2)
    max_page = int(driver.find_element('class name', "page").text.split(' / ')[-1])
    url = []
    for i in range(1, max_page+1):
        driver.execute_script(f"list({i});")
        url += [li.find_element('tag name', 'a').get_attribute('href') for li in driver.find_elements('class name', 'ptitle1')]

    for u in url:
        driver.get(u)
        try:
            driver.get(driver.find_element('xpath', "//a[contains(@title, '生效案件.xls')]").get_attribute('href'))
        except:
            pass
    # %%
    file_list = glob('tmp/預計*')

    p = Pool(int(cpu_count() * 0.8))
    result = list(tqdm(p.imap(file_process, file_list), total=len(file_list)))
    p.close()
    p.join()
    
    p = Pool(int(cpu_count() * 0.8))
    result2 = list(tqdm(p.imap(data_process, result), total=len(result)))
    p.close()
    p.join()
    
    result = result2