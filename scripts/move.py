import os
import shutil
from dotenv import load_dotenv

load_dotenv()

dest_folder = os.getenv('DEST_FOLDER_MOVE')

if not os.path.exists(dest_folder):
    os.makedirs(dest_folder)

root_folder = os.getenv('ROOT_FOLDER_MOVE')

for root, dirs, files in os.walk(root_folder):
    for file in files:
        if file.endswith('.xml'):
            src_file = os.path.join(root, file)

            folder_name = os.path.basename(root)
            dest_file = os.path.join(dest_folder, f"{folder_name}.xml")

            shutil.copy(src_file, dest_file)
            print(f'Arquivo {src_file} copiado para {dest_file}')

print('Fim do programa')
