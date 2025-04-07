from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd
import time
from selenium.webdriver.firefox.options import Options as foptions
from selenium.webdriver.edge.options import Options as eoptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from lxml import etree

uni_name = []
sehir=[]
logo_url = []
banner_url = []
type = []

Foptions = foptions()
Foptions.headless = False

Eoptions = eoptions()
Eoptions.headless = False

links = ["https://yokatlas.yok.gov.tr/universite.php"]


for link in links:
    driver = webdriver.Firefox(options=Foptions)
    driver.get(link)
    driver.maximize_window()
    time.sleep(7)
    
    html = driver.page_source
    soup = BeautifulSoup(html,"html.parser")
    WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="myUl"]'))
        )
    list = driver.find_element(By.XPATH, ('//*[@id="myUl"]'))
    list_items = list.find_elements(By.TAG_NAME, 'li')

    # Get the length
    length = len(list_items)
    
    html = driver.page_source
    soup = BeautifulSoup(html,"html.parser")

    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '/html/body'))
    )

    for x in range(1,length+1):

        try:
            name_var = driver.find_element(By.XPATH,(f'//*[@id="myUl"]/li[{x}]/div/div[1]/h3')).text
            uni_name.append(name_var)
        except:
            uni_name.append(None)

        try:
            WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, f'//*[@id="myUl"]/li[{x}]/div/a/img'))
        )
            logo = driver.find_element(By.XPATH,(f'//*[@id="myUl"]/li[{x}]/div/a/img')).get_attribute("src")
            logo_url.append(logo)
        except:
            logo_url.append(None)
        
        try:
            sehir_var = driver.find_element(By.XPATH,(f'//*[@id="myUl"]/li[{x}]/div/div[1]/span[2]')).text
            sehir.append(sehir_var)
        except:
            sehir.append(None)

        try:
            type_var =  driver.find_element(By.XPATH,(f'//*[@id="myUl"]/li[{x}]/div/div[1]/span[1]')).text
            type.append(type_var)
        except:
            type.append(None)

        banner_url.append(None)
    

    driver.quit()

uni_list = pd.DataFrame()

uni_list.insert(0,"Üniversite İsmi","")
uni_list.insert(1,"sehir","")
uni_list.insert(2,"logo","")
uni_list.insert(3,"banner","")
uni_list.insert(4,"üniversite Türü","")
for uni in range(0,len(uni_name)):
        uni_info = {"Üniversite İsmi":uni_name[uni],
                    "sehir": sehir[uni],
                    "logo": logo_url[uni],
                    "banner": banner_url[uni],
                    "üniversite Türü": type[uni]
                    }
        uni_list = uni_list.append(uni_info, ignore_index=True)

uni_list.to_excel("unis_start.xlsx",index=True,merge_cells=False)
