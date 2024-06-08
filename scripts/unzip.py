import os
import zipfile

# Subistitua o valor da variável compressed_folder pelo caminho da pasta onde estão os arquivos compactados
compressed_folder = 'C:/Users/PedroDL/Documents/ProdMais Insper/ProdMais-Utils/compactadas'

# Subistitua o valor da variável dest_folder pelo caminho da pasta onde deseja descompactar os arquivos
dest_folder = 'C:/Users/PedroDL/Documents/ProdMais Insper/ProdMais-Utils/descompactadas'

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