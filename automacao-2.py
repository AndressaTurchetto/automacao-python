from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

def preencher_campo(driver, campo, valor, descricao=""):
    try:
        driver.execute_script("arguments[0].value = '';", campo)
        time.sleep(0.5)
        
        driver.execute_script(f"arguments[0].value = '{valor}';", campo)
        time.sleep(0.5)
        
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", campo)
        time.sleep(0.5)
                
        print(f"Campo {descricao} preenchido com: {valor}")
        return True
    except Exception as e:
        print(f"Erro ao preencher {descricao}: {str(e)}")
        return False

def automatizar_dnit():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", {
        "download.default_directory": r"C:\chromedriver\downloads",
        "download.prompt_for_download": False
    })
    
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)
    
    try:
        driver.get("https://servicos.dnit.gov.br/vgeo/")
        print("Site acessado")
        time.sleep(5)
        
        print("Procurando botão togglegeo...")
        togglegeo = wait.until(EC.element_to_be_clickable((By.ID, "togglegeo")))
        togglegeo.click()
        print("Togglegeo clicado")
        time.sleep(2)
        
        print("Procurando botão dgept...")
        dgept = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "dgept.geobtn.gbtnopt")))
        dgept.click()
        print("Dgept clicado")
        time.sleep(2)
        
        df = pd.read_excel(r"C:\chromedriver\planilha.xlsx")
        print(f"Planilha lida com {len(df)} linhas")
        
        for index, row in df.iterrows():
            print(f"\nProcessando linha {index + 1}")
            try:
                time.sleep(2)
                campos = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "infoform")))
                
                preencher_campo(driver, campos[0], "PR", "UF")
                preencher_campo(driver, campos[1], "277", "BR")
                preencher_campo(driver, campos[2], "B", "TRECHO")
                
                preencher_campo(driver, campos[3], str(row['KM']), "KM")
                
                print("Procurando botão gesppointbtn...")
                gesppointbtn = wait.until(EC.element_to_be_clickable((By.ID, "gesppointbtn")))
                gesppointbtn.click()
                print("Gesppointbtn clicado")
                time.sleep(2)
                
                print("Procurando botão dglayer...")
                dglayer = wait.until(EC.element_to_be_clickable(
                    (By.CLASS_NAME, "dglayer.geobtn.gbtnopt.gbtnactive")))
                dglayer.click()
                print("Dglayer clicado")
                time.sleep(2)
                
                print("Procurando botão de download...")
                faddbtn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "faddbtn")))
                faddbtn.click()
                print("Download iniciado")
                
                time.sleep(5)
                print(f"Linha {index + 1} processada com sucesso")
                
            except Exception as e:
                print(f"Erro ao processar linha {index + 1}: {str(e)}")
                driver.save_screenshot(f"erro_linha_{index + 1}.png")
                
                driver.refresh()
                time.sleep(5)
                
                togglegeo = wait.until(EC.element_to_be_clickable((By.ID, "togglegeo")))
                togglegeo.click()
                time.sleep(2)
                
                dgept = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "dgept.geobtn.gbtnopt")))
                dgept.click()
                time.sleep(2)
                
                continue
    
    except Exception as e:
        print(f"Erro geral: {str(e)}")
        driver.save_screenshot("erro_geral.png")
    
    finally:
        print("\nFinalizando automação...")
        time.sleep(3)
        driver.quit()

if __name__ == "__main__":
    automatizar_dnit()