from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from dotenv import load_dotenv
import os
import openpyxl

load_dotenv()

browser = webdriver.Firefox()
# browser = webdriver.Firefox(executable_path='C:\\geckodriver.exe')

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
# cvs_dir = 'C:\\Users\\Yan\\Desktop\\utilsProdmais\\ProdMais-Utils\\curriculos'

c = 1
# local dos docentes é o mesmo do arquivo atual + dados_docentes.csv
dados_docentes_dir = 'data\\dados_docentes.xlsx'

# Carregue o arquivo de workbook apenas uma vez
workbook = openpyxl.load_workbook(dados_docentes_dir)
sheet = workbook['docentes_lattes']

for row in range(2, sheet.max_row + 1):
    url_lattes = sheet.cell(row=row, column=3).value
    id = url_lattes.split('/')[-1]
    print(id)

    for filename in os.listdir(cvs_dir):
        if filename.endswith('.xml') and id in filename:
            file_path = os.path.join(cvs_dir, filename)

            browser.find_element(By.ID, 'fileXML').send_keys(file_path)
            time.sleep(1)

            # encontra campo name="genero" e digita o conteúdo da coluna 12
            genero = sheet.cell(row=row, column=12).value
            browser.find_element(By.NAME, 'genero').clear()
            browser.find_element(By.NAME, 'genero').send_keys(genero)
            time.sleep(0.5)
            
            browser.find_element(By.XPATH, '/html/body/main/div/div/form[1]/div[2]/button').click()
            time.sleep(2)

            page_content = browser.page_source
            if ('Registro anterior não encontrado na base') in page_content:
                print(f"Arquivo {filename} [{c}] adicionado")
                sheet.cell(row=row, column=9).value = 'OK'
                sheet.cell(row=row, column=10).value = time.strftime('%H:%M:%S %d/%m/%Y')
            else:
                print(f"Arquivo {filename} [{c}] já existe na base")
                sheet.cell(row=row, column=9).value = 'Já existe'
                sheet.cell(row=row, column=10).value = time.strftime('%H:%M:%S %d/%m/%Y')
            c+=1

            browser.back()
            time.sleep(2)

            # Salve as alterações a cada iteração
            workbook.save(dados_docentes_dir)

end_time = time.time()
final_time = (end_time - start_time)/60

print(f"Tempo de execução do script: {final_time:2f} minutos")

browser.quit()
