import os
import shutil

dest_folder = 'C:/Users/PedroDL/Documents/ProdMais Insper/ProdMais-Utils/curriculos'

if not os.path.exists(dest_folder):
    os.makedirs(dest_folder)

root_folder = 'C:/Users/PedroDL/Documents/ProdMais Insper/ProdMais-Utils/descompactadas'

for root, dirs, files in os.walk(root_folder):
    for file in files:
        if file.endswith('.xml'):
            src_file = os.path.join(root, file)

            folder_name = os.path.basename(root)
            dest_file = os.path.join(dest_folder, f"{folder_name}.xml")

            shutil.move(src_file, dest_file)
            print(f'Arquivo {src_file} copiado para {dest_file}')

print('Fim do programa')
