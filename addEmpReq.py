import requests
import openpyxl
import random


numeros_possiveis = list(range(60000, 65001))

# Função para obter um ID aleatório caso o do cliente esteja em uso
def obter_id_aleatorio():
    if numeros_possiveis:  # Verifica se ainda há números disponíveis
        return numeros_possiveis.pop(random.randint(0, len(numeros_possiveis) - 1))
    else:
        raise ValueError("Todos os IDs possíveis já foram usados.")


#SecretKey
secret_key = '9561fd545cf7e0836b60468936592ec2'  # Substitua pelo seu token real

# Carrega a planilha
workbook = openpyxl.load_workbook('empresa.xlsx')
sheet = workbook['Empresas']

for row in sheet.iter_rows(min_row=5, min_col=1, max_col=11):
    id = row[0].value          # Coluna 1: Código do Sistema Contábil (ID)
    fantasia = row[1].value       # Coluna 2: Nome Fantasia
    nome = row[2].value        # Coluna 3: Razão Social
    cnpj = row[3].value               # Coluna 4: CNPJ/CPF/CEI/CAEPF
    InscEst = row[4].value            # Coluna 5: Insc. Estadual
    InscUf = row[5].value             # Coluna 6: Insc. Estadual UF
    NomeCtto = row[6].value           # Coluna 7: Nome do Contato
    EmailCtto = row[7].value          # Coluna 8: E-mail do Contato
    Regime = row[8].value             # Coluna 9: Regime Tributário
    Apelido = row[9].value           # Coluna 11: Apelido e-Contínuo




    # Construindo a URL com parâmetros de consulta
    url = f'https://api.acessorias.com/companies'
    params = {
        'cnpj': cnpj,
        'nome': nome,
        'fantasia' : fantasia,
        'id' : id,
    }

    # Headers para a requisição, incluindo o token de autorização
    headers = {
        'Authorization': f'Bearer {secret_key}',
        'Content-Type': 'application/json'
    }

    # Enviando a requisição GET
    response = requests.post(url, headers=headers, params=params)

    # Imprimindo o código de status e a resposta JSON
    print(f"Status Code: {response.status_code}")
    try:
        # Tentando decodificar a resposta como JSON
        response_json = response.json()
        print("Resposta JSON:", response_json)
    except requests.exceptions.JSONDecodeError:
        # Caso a resposta não seja JSON, imprime o conteúdo da resposta como texto
        print("Resposta não é um JSON.")
        print("Resposta:", response.text)
        
    teste = input('breakpoint')