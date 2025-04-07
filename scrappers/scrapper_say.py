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
fakulte=[]
bolum = []
yop = []
burs = []
type = []
doluluk = []
puan = []
son4_kont = []
son4_yerleşen= []
son4_sıralama = []
son4_puan = []
bolum_url= []

Foptions = foptions()
Foptions.headless = True

Eoptions = eoptions()
Eoptions.headless = True

links = ["https://yokatlas.yok.gov.tr/tercih-sihirbazi-t4-tablo.php?p=say"]


for link in links:
    driver = webdriver.Firefox(options=Foptions)
    driver.get(link)
    driver.maximize_window()
    time.sleep(7)
    
    html = driver.page_source
    soup = BeautifulSoup(html,"html.parser")


    for page in range(1,110):
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="mydata"]/tbody'))
        )
        list = driver.find_element(By.XPATH, ('//*[@id="mydata"]/tbody'))
        list_items = list.find_elements(By.TAG_NAME, 'tr')

        # Get the length
        length = len(list_items)
        
        html = driver.page_source
        soup = BeautifulSoup(html,"html.parser")

        for x in range(1,length+1):

            try:
                name_var = driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[3]/strong')).text
                uni_name.append(name_var)
            except:
                uni_name.append(None)
            
            try:
                fakulte_var = driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[3]/font')).text
                fakulte.append(fakulte_var)
            except:
                fakulte.append(None)

            try:
                bolum_var = driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[4]/strong')).text
                bolum.append(bolum_var)
            except:
                bolum.append(None)

            try:
                yop_var = driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[2]/a[1]')).text
                yop.append(yop_var)
            except:
                yop.append(None)

            try:
                bolum_url_var = driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[2]/a[1]')).get_attribute('href')
                bolum_url.append(bolum_url_var)
            except:
                bolum_url.append(None)

            try:
                burs_var = driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[7]')).text
                burs.append(burs_var)
            except:
                burs.append(None)

            try:
                type_var =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[8]')).text
                type.append(type_var)
            except:
                type.append(None)

            try:
                doluluk_var =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[10]')).text
                doluluk.append(doluluk_var)
            except:
                doluluk.append(None)

            puan.append('SAY')

            try:
                son4_kont_var_1 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[9]/font[1]')).text
                son4_kont_var_2 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[9]/font[2]')).text
                son4_kont_var_3 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[9]/font[3]')).text
                son4_kont_var_4 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[9]/font[4]')).text
                try:
                    if son4_kont_var_1 == "---":
                        son4_kont_var_1= "0+0+0+0"
                    if son4_kont_var_2 == "---":
                        son4_kont_var_2= "0+0+0+0"
                    if son4_kont_var_3 == "---":
                        son4_kont_var_3= "0+0+0+0"
                    if son4_kont_var_4 == "---":
                        son4_kont_var_4= "0+0+0+0"
                except:
                    pass

                asıl = int(son4_kont_var_1.split('+')[0]) + int(son4_kont_var_2.split('+')[0]) + int(son4_kont_var_3.split('+')[0]) + int(son4_kont_var_4.split('+')[0])
                ek = int(son4_kont_var_1.split('+')[1]) + int(son4_kont_var_2.split('+')[1]) + int(son4_kont_var_3.split('+')[1]) + int(son4_kont_var_4.split('+')[1]) + int(son4_kont_var_1.split('+')[2]) + int(son4_kont_var_2.split('+')[2]) + int(son4_kont_var_1.split('+')[3]) + int(son4_kont_var_2.split('+')[3])
                son4_kont.append(str(asıl) + '+' + str(ek))
            except:
                son4_kont.append(None)

            

            try:
                son4_yerleşen_var_1 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[11]/font[1]')).text
                son4_yerleşen_var_2 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[11]/font[2]')).text
                son4_yerleşen_var_3 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[11]/font[3]')).text
                son4_yerleşen_var_4 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[11]/font[4]')).text

                temp_list = [son4_yerleşen_var_1, son4_yerleşen_var_2, son4_yerleşen_var_3, son4_yerleşen_var_4]

                son4_yerleşen.append(temp_list)
            except:
                son4_yerleşen.append(None)

            try:
                son4_sıralama_var_1 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[12]/font[1]')).text
                son4_sıralama_var_2 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[12]/font[2]')).text
                son4_sıralama_var_3 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[12]/font[3]')).text
                son4_sıralama_var_4 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[12]/font[4]')).text

                temp_list = [son4_sıralama_var_1, son4_sıralama_var_2, son4_sıralama_var_3, son4_sıralama_var_4]

                son4_sıralama.append(temp_list)
            except:
                son4_sıralama.append(None)

            try:
                son4_puan_var_1 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[13]/font[1]')).text
                son4_puan_var_2 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[13]/font[2]')).text
                son4_puan_var_3 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[13]/font[3]')).text
                son4_puan_var_4 =  driver.find_element(By.XPATH,(f'//*[@id="mydata"]/tbody/tr[{x}]/td[13]/font[4]')).text

                temp_list = [son4_puan_var_1, son4_puan_var_2, son4_puan_var_3, son4_puan_var_4]

                son4_puan.append(temp_list)
            except:
                son4_puan.append(None)

        next_page = driver.find_element(By.XPATH, '//*[@id="mydata_next"]/a')
        next_page.click()
        time.sleep(2)
    driver.quit()

uni_list = pd.DataFrame()


uni_list.insert(0,"Üniversite İsmi","")
uni_list.insert(1,"fakülte","")
uni_list.insert(2,"bolum","")
uni_list.insert(3,"yop","")
uni_list.insert(4,"burs","")
uni_list.insert(5,"type","")
uni_list.insert(6,"doluluk","")
uni_list.insert(7,"puan","")
uni_list.insert(8,"son4_kont","")
uni_list.insert(9,"son4_yerleşen","")
uni_list.insert(10,"son4_sıralama","")
uni_list.insert(11,"son4_puan","")
uni_list.insert(12,"bolum_url","")
for uni in range(0,len(uni_name)):
        uni_info = {"Üniversite İsmi":uni_name[uni],
                "fakülte":fakulte[uni],
                "bolum":bolum[uni],
                "yop":yop[uni],
                "burs":burs[uni],
                "type":type[uni],
                "doluluk":doluluk[uni],
                "puan":puan[uni],
                "son4_kont":son4_kont[uni],
                "son4_yerleşen":son4_yerleşen[uni],
                "son4_sıralama":son4_sıralama[uni],
                "son4_puan":son4_puan[uni],
                "bolum_url":bolum_url[uni]
                }
        uni_list = uni_list.append(uni_info, ignore_index=True)

uni_list.to_excel("unis_details_say.xlsx",index=True,merge_cells=False)
