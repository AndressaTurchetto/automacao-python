from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import pandas as pd
import time
import os

def esperar_e_clicar(driver, by, valor, tempo=20):
    """Função auxiliar para esperar elemento e clicar com retry"""
    wait = WebDriverWait(driver, tempo)
    for tentativa in range(3):  # Tenta 3 vezes
        try:
            elemento = wait.until(EC.element_to_be_clickable((by, valor)))
            time.sleep(1)  # Pequena pausa para garantir que o elemento está realmente pronto
            elemento.click()
            return elemento
        except ElementClickInterceptedException:
            time.sleep(2)  # Espera um pouco mais se o elemento estiver interceptado
            continue
        except Exception as e:
            print(f"Tentativa {tentativa + 1} falhou: {str(e)}")
            if tentativa == 2:  # Na última tentativa, relança o erro
                raise
    return None

def esperar_e_preencher(driver, by, valor, texto, tempo=20):
    """Função auxiliar para esperar elemento e preencher com retry"""
    wait = WebDriverWait(driver, tempo)
    for tentativa in range(3):  # Tenta 3 vezes
        try:
            elemento = wait.until(EC.presence_of_element_located((by, valor)))
            time.sleep(1)  # Pequena pausa para garantir que o elemento está realmente pronto
            elemento.clear()
            elemento.send_keys(Keys.CONTROL + "a")  # Seleciona todo o texto
            elemento.send_keys(Keys.DELETE)  # Deleta o texto selecionado
            elemento.send_keys(str(texto))
            return elemento
        except Exception as e:
            print(f"Tentativa {tentativa + 1} falhou: {str(e)}")
            if tentativa == 2:  # Na última tentativa, relança o erro
                raise
    return None

def automatizar_dnit():
    # Configurar o Chrome Driver
    options = webdriver.ChromeOptions()
    # Definir pasta de download padrão
    prefs = {
        "download.default_directory": r"C:\chromedriver\downloads",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    
    try:
        # Acessar o site
        driver.get("https://servicos.dnit.gov.br/vgeo/")
        print("Site acessado com sucesso")
        
        # Esperar a página carregar completamente
        time.sleep(5)
        
        # Clicar no togglegeo
        print("Tentando clicar no togglegeo...")
        esperar_e_clicar(driver, By.ID, "togglegeo")
        print("Clique no togglegeo realizado")
        
        # Clicar no botão dgept
        print("Tentando clicar no botão dgept...")
        time.sleep(2)  # Espera adicional para garantir que o menu esteja visível
        esperar_e_clicar(driver, By.CLASS_NAME, "dgept.geobtn.gbtnopt")
        print("Clique no dgept realizado")
        
        # Ler a planilha
        try:
            df = pd.read_excel(r"C:\chromedriver\planilha.xlsx")
            print("Planilha lida com sucesso")
        except Exception as e:
            print(f"Erro ao ler planilha: {e}")
            return
        
        # Para cada linha na planilha
        for index, row in df.iterrows():
            print(f"\nProcessando linha {index + 1}")
            try:
                # Esperar todos os campos estarem presentes
                time.sleep(2)
                campos = WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, "infoform")))
                
                # Preencher campos
                print("Preenchendo UF...")
                esperar_e_preencher(driver, By.CLASS_NAME, "infoform", row['UF'])
                
                print("Preenchendo BR...")
                campos[1].send_keys(Keys.CONTROL + "a")
                campos[1].send_keys(Keys.DELETE)
                campos[1].send_keys(str(row['BR']))
                
                print("Preenchendo Trecho...")
                campos[2].send_keys(Keys.CONTROL + "a")
                campos[2].send_keys(Keys.DELETE)
                campos[2].send_keys(row['TRECHO'])
                
                print("Preenchendo KM...")
                campos[3].send_keys(Keys.CONTROL + "a")
                campos[3].send_keys(Keys.DELETE)
                campos[3].send_keys(str(row['KM']))
                
                # Clicar nos botões necessários
                print("Clicando no gesppointbtn...")
                esperar_e_clicar(driver, By.ID, "gesppointbtn")
                
                time.sleep(3)  # Espera adicional para o carregamento
                
                print("Clicando no dglayer...")
                esperar_e_clicar(driver, By.CLASS_NAME, "dglayer.geobtn.gbtnopt.gbtnactive")
                
                print("Clicando no botão de download...")
                esperar_e_clicar(driver, By.CLASS_NAME, "faddbtn")
                
                # Aguardar download completar
                time.sleep(5)
                print(f"Linha {index + 1} processada com sucesso")
                
            except Exception as e:
                print(f"Erro ao processar linha {index + 1}: {e}")
                continue
    
    except Exception as e:
        print(f"Ocorreu um erro geral: {e}")
    
    finally:
        print("\nFinalizando automação...")
        time.sleep(3)  # Espera final para garantir que tudo foi concluído
        driver.quit()

if __name__ == "__main__":
    automatizar_dnit()