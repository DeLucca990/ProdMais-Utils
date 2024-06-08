import pyautogui
import time
import webbrowser

def get_mouse_position():
    print("Mova o mouse para a posição desejada e espere 5 segundos...")
    time.sleep(5)
    x, y = pyautogui.position()
    print(f"A posição do mouse é: ({x}, {y})")

    with open("./assets/mouse_position.txt", "w") as file:
        file.write(f"{x},{y}")    
    print("A posição do mouse foi salva no arquivo 'mouse_position.txt'.")

    return x, y

def main():
    url = "http://buscatextual.cnpq.br/buscatextual/download.do?metodo=apresentar&idcnpq=6758933610036027"
    firefox_path = 'C:/Arquivos de Programas/Mozilla Firefox/firefox.exe'

    webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(firefox_path))

    webbrowser.get('firefox').open(url)

    get_mouse_position()
    
if __name__ == "__main__":
    main()