import pyautogui
import time
import webbrowser
import os
import pandas as pd

def read_mouse_position():
    try:
        with open("./assets/mouse_position.txt", "r") as file:
            position = file.read().strip()
            x, y = map(int, position.split(','))
            print(f"A posição do mouse é: ({x}, {y})")
            return x, y
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None

def is_downloaded(download_folder, id, timeout=60, check_interval=5):
    elapsed_time = 0    
    while elapsed_time < timeout:
        current_files = set(os.listdir(download_folder))
        for file in current_files:
            # verifica se existe um id no nome do arquivo
            if id in file and file.endswith('.zip'):
                return file
        time.sleep(check_interval)
        elapsed_time += check_interval

    print("Nenhum arquivo novo foi detectado.")
    return None

def extract_id_lattes(url):
    return url.split('/')[-1]

def download_lattes(id_lattes, x_position, y_position, download_folder):
    url = f"http://buscatextual.cnpq.br/buscatextual/download.do?metodo=apresentar&idcnpq={id_lattes}"
    firefox_path = 'C:/Arquivos de Programas/Mozilla Firefox/firefox.exe'  # Caminho do executável do Firefox

    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(firefox_path))

    try:
        while True:
            print(f"Attempting to open URL: {url}")
            webbrowser.get('firefox').open(url)
            print("URL aberta com sucesso.")

            # Esperar 15 segundos para carregar a extensão do anticaptcha
            time.sleep(15)

            # Verifica se o download foi completado
            downloaded_file = is_downloaded(download_folder, id_lattes)
            if downloaded_file:
                print(f"O arquivo {id_lattes} foi baixado com sucesso como {downloaded_file}.")
                break
            else:
                print(f"O download {id_lattes} não foi detectado. Tentando novamente...")

            # Fechar a aba atual do navegador
            pyautogui.hotkey('ctrl', 'w')
        
        pyautogui.hotkey('ctrl', 'w')
        return True
    except KeyboardInterrupt:
        print("Processo interrompido.")
        return False

def main():
    start_time = time.time()
    df = pd.read_excel("./data/dados_docentes_prodmais.xlsx", sheet_name="docentes_lattes") 
    download_folder = 'C:\\Users\\Yan\\Downloads'  # Caminho explícito para a pasta de downloads
    x_position, y_position = read_mouse_position()

    for index, url_lattes in enumerate(df['ds_url_lattes']):
        id_lattes = extract_id_lattes(url_lattes)
        print('\n')
        print(f"Iniciando download para ID: {id_lattes}")
        print(f"Indice: {index}")
        if download_lattes(id_lattes, x_position, y_position, download_folder):
            print(f"Download para ID: {id_lattes} concluído com sucesso.")
        else:
            print(f"Download para ID: {id_lattes} falhou.")
            break
    print("Salvando arquivo...")

    end_time = time.time()
    final_time = (end_time - start_time) / 60
    print(f"Tempo total de execução: {final_time:.2f} minutos.")

if __name__ == "__main__":
    main()
