import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from dotenv import load_dotenv
import os

load_dotenv()

browser = webdriver.Firefox()

url = 'http://localhost:8080/inclusao.php'
username = os.getenv('PRODMAIS_USERNAME')
password = os.getenv('PRODMAIS_PASSWORD')

browser.get(url)

start_time = time.time()

browser.find_element(By.NAME, 'username').send_keys(username)
browser.find_element(By.NAME, 'password').send_keys(password)
time.sleep(0.5)
browser.find_element(By.NAME, 'submit').click()
time.sleep(1)

cvs_dir = '/home/pedrodl/Documents/ProdMaisInsper/ProdMais-Utils/curriculos'
c = 1
for filename in os.listdir(cvs_dir):
    if filename.endswith('.xml'):
        file_path = os.path.join(cvs_dir, filename)

        browser.find_element(By.ID, 'fileXML').send_keys(file_path)
        time.sleep(1)

        browser.find_element(By.XPATH, '/html/body/main/div/div/form[1]/div[2]/button').click()
        time.sleep(2)

        # Verificando conteúdo da página
        page_content = browser.page_source
        if ('Registro anterior não encontrado na base') in page_content:
            print(f"Arquivo {filename} [{c}] adicionao")
        else:
            print(f"Arquivo {filename} [{c}] já existe na base")
        c+=1

        browser.back()
        time.sleep(2)

end_time = time.time()
final_time = (end_time - start_time)/60

print(f"Tempo de execução do script: {final_time:2f} minutos")

browser.quit()