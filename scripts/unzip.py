import os
import zipfile
from dotenv import load_dotenv

load_dotenv()

compressed_folder = os.getenv('COMPRESSED_FOLDER_UNZIP')

dest_folder = os.getenv('DEST_FOLDER_UNZIP')

if not os.path.exists(dest_folder):
    os.makedirs(dest_folder)

for root, dirs, files in os.walk(compressed_folder):
    for file in files:
        if file.endswith('.zip'):
            zip_file_path = os.path.join(root, file)
            extract_folder = os.path.join(dest_folder, os.path.splitext(file)[0])
            if not os.path.exists(extract_folder):
                os.makedirs(extract_folder)
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_folder)
            print(f'Descompactado: {zip_file_path} para {extract_folder}')

print('Fim de programa.')