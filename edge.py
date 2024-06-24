import os
import pandas as pd
from time import sleep
from multiprocessing import Pool, cpu_count
from glob import glob

from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.microsoft import EdgeChromiumDriverManager


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


if __name__ == '__main__':
    if not os.path.exists('tmp'):
        os.makedirs('tmp')

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