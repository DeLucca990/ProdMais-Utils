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

def is_downloaded(download_folder, wait_time=5):
    before = set(os.listdir(download_folder))
    time.sleep(wait_time)
    after = set(os.listdir(download_folder))
    downloaded_files = after - before
    if downloaded_files:
        print(f"Download completado: {downloaded_files}")
        return True
    else:
        print("Nenhum arquivo novo foi detectado.")
        return False

def download_lattes(id_lattes, x_position, y_position, download_folder):
    url = f"http://buscatextual.cnpq.br/buscatextual/download.do?metodo=apresentar&idcnpq={id_lattes}"
    firefox_path = 'C:/Arquivos de Programas/Mozilla Firefox/firefox.exe'  # Caminho do executável do Firefox

    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(firefox_path))

    try:
        while True:
            webbrowser.get('firefox').open(url)
            print("Abrindo URL...")

            # Esperar 10 segundos para carregar a extensão do anticaptcha
            time.sleep(10)

            # Mova o mouse para a posição capturada e clique
            pyautogui.moveTo(x_position, y_position)
            time.sleep(1.5)
            pyautogui.click()

            # Verifica se o download foi completado
            if is_downloaded(download_folder):
                print(f"O arquivo {id_lattes} foi baixado com sucesso.")
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
    df = pd.read_excel("./data/dados docentes prodmais.xlsx", sheet_name="docentes_lattes") 
    download_folder = os.path.expanduser('~\\Downloads') 
    x_position, y_position = read_mouse_position()

    for index, id_lattes in enumerate(df['id_lattes']):
        # if index < 300:
        #     continue
        print('\n')
        print(f"Iniciando download para ID: {id_lattes}")
        print(f"Indice: {index}")
        if download_lattes(id_lattes, x_position, y_position, download_folder):
            df.loc[index, 'status_extracao'] = 1
            print(f"Download para ID: {id_lattes} concluído com sucesso.")
        else:
            df.loc[index, 'status_extracao'] = 0
            print(f"Download para ID: {id_lattes} falhou.")
            break
        # if index == 299:
        #     break
    print("Salvando arquivo...")
    df.to_excel("./data/resultado.xlsx", index=False)

    end_time = time.time()
    final_time = (end_time - start_time) / 60
    print(f"Tempo total de execução: {final_time:.2f} minutos.")

if __name__ == "__main__":
    main()
