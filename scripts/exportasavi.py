import time
import requests
from dotenv import load_dotenv
import os
import pandas as pd
load_dotenv()

# url = 'http://localhost:8080/inclusao.php'
url = 'https://prodmais.datascience.insper.edu.br/inclusao.php'
username = os.getenv('ELASTIC_USERNAME')
password = os.getenv('ELASTIC_PASSWORD')

# URL da sua API de relatórios
url = 'https://prodmaiselastic.datascience.insper.edu.br/api/reporting/generate/csv_searchsource?jobParams=%28browserTimezone%3AAmerica%2FSao_Paulo%2Ccolumns%3A%21%28tipo%2Cvinculo.lattes_id%2CdatePublished%2Cvinculo.nome%2Cauthor.person%2Cauthor.ordemDeAutoria%2Cname%2CalternateName%2Cdoi%2Cabout%2Clanguage%2Cauthor.nomeParaCitacao%2C_score%2CisPartOf.name%2Clattes.natureza%2CisPartOf.issn%29%2CobjectType%3Asearch%2CsearchSource%3A%28fields%3A%21%28%28field%3Atipo%2Cinclude_unmapped%3Atrue%29%2C%28field%3Avinculo.lattes_id%2Cinclude_unmapped%3Atrue%29%2C%28field%3AdatePublished%2Cinclude_unmapped%3Atrue%29%2C%28field%3Avinculo.nome%2Cinclude_unmapped%3Atrue%29%2C%28field%3A%27author.person.%2A%27%2Cinclude_unmapped%3Atrue%29%2C%28field%3Aauthor.ordemDeAutoria%2Cinclude_unmapped%3Atrue%29%2C%28field%3Aname%2Cinclude_unmapped%3Atrue%29%2C%28field%3AalternateName%2Cinclude_unmapped%3Atrue%29%2C%28field%3Adoi%2Cinclude_unmapped%3Atrue%29%2C%28field%3Aabout%2Cinclude_unmapped%3Atrue%29%2C%28field%3Alanguage%2Cinclude_unmapped%3Atrue%29%2C%28field%3Aauthor.nomeParaCitacao%2Cinclude_unmapped%3Atrue%29%2C%28field%3A_score%2Cinclude_unmapped%3Atrue%29%2C%28field%3AisPartOf.name%2Cinclude_unmapped%3Atrue%29%2C%28field%3Alattes.natureza%2Cinclude_unmapped%3Atrue%29%2C%28field%3AisPartOf.issn%2Cinclude_unmapped%3Atrue%29%29%2Cfilter%3A%21%28%29%2Cindex%3A%27221e1c1e-7b44-4339-8354-d70747c8a821%27%2Cquery%3A%28language%3Akuery%2Cquery%3A%27%27%29%2Csort%3A%21%28%28_score%3Adesc%29%29%29%2Ctitle%3AExporta%C3%A7%C3%A3oSavi%2Cversion%3A%278.13.4%27%29'
# Dados de autenticação, se necessário
auth = (username, password)  # Substitua com suas credenciais

# Headers necessários
headers = {
    'kbn-xsrf': 'true'  # Substitua com o valor apropriado do kbn-xsrf
}

# Função para verificar o status do job
def check_job_status(job_path):
    status_url = f"https://prodmaiselastic.datascience.insper.edu.br{job_path}"
    while True:
        response = requests.get(status_url, headers=headers, auth=auth, verify=False)
        if response.status_code == 200:
            return response
        elif response.status_code == 503:
            print('Job ainda pendente, aguardando...')
            time.sleep(10)  # Aguarda 10 segundos antes de verificar novamente
        else:
            print(f'Erro no download {response.status_code}: {response.text}')
            return None

# Realiza a requisição POST
response = requests.post(url, headers=headers, auth=auth, verify=False)

# Verifica o status da resposta
if response.status_code == 200:
    # Sucesso, imprime o conteúdo da resposta (geralmente um JSON com informações)
    print(response.json())
    job_path = response.json().get('path')
    
    if job_path:
        # Verifica o status do job e baixa o conteúdo quando estiver pronto
        download_response = check_job_status(job_path)
        
        if download_response:
            # Salva o conteúdo em um arquivo
            with open('report.csv', 'wb') as f:
                f.write(download_response.content)
            print('Download completo e salvo como report.csv')
            
            # Lê o arquivo CSV
            df = pd.read_csv('report.csv')

            # Converte para Excel
            df.to_excel('report.xlsx', index=False)

            print('CSV convertido para Excel com sucesso!')
    else:
        print('Job path não encontrado na resposta')
else:
    # Erro, imprime o status code e a mensagem de erro
    print(f'Erro {response.status_code}: {response.text}')
