import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

df = pd.read_excel(r'C:\chromedriver\planilha.xlsx')

service = Service(r'C:\chromedriver\chromedriver.exe')
driver = webdriver.Chrome(service=service)
driver.get("https://servicos.dnit.gov.br/vgeo/")

WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "togglegeo"))).click()

WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Espacializar ponto SNV']"))).click()

WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "gepuf")))

for index, row in df.iterrows():

    uf_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "gepuf")))
    uf_input.send_keys(row['UF'])

    br_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "gepbr")))
    br_input.send_keys("277")

    br_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "geptptr")))
    br_input.send_keys("B")

    km_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "gepkm")))

    if 'KM' in row:
        km_value = row['KM'] 
        km_input.clear()
        km_input.send_keys(km_value)

        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "gesppointbtn"))).click()
    else:
        print("Coluna 'KM' não encontrada na linha atual.")

driver.quit()
