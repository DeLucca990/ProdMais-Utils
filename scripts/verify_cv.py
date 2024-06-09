import pandas as pd
import os
import xml.etree.ElementTree as ET

def check_error_xml(file_path):
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        mensagem_element = root.find('.//MENSAGEM')
        if mensagem_element is not None and mensagem_element.text == 'Erro ao recuperar o XML':
            return True
    except ET.ParseError:
        return False
    return False

cvs_folder_path = 'C:/Users/PedroDL/Documents/ProdMais Insper/ProdMais-Utils/curriculos'
df = pd.read_excel("./data/dados docentes prodmais.xlsx", sheet_name="docentes_lattes") 

for index, row in df.iterrows():
    xml_file_path = os.path.join(cvs_folder_path, f"{row['id_lattes']}.xml")
    if os.path.exists(xml_file_path):
        if check_error_xml(xml_file_path):
            df.at[index, 'status_cv'] = 'Erro'
        else:
            df.at[index, 'status_cv'] = 'Ok'
    else:
        df.at[index, 'status_cv'] = 'Arquivo não encontrado'

print(df)

df.to_excel("./data/resultado.xlsx", index=False)