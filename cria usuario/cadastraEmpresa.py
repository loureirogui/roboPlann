import traceback
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from docx import Document
import openpyxl
import unicodedata

# Solicita as credenciais do usuário
print("Cadastrando Empresas Agora")

emailLogin = 'guilherme.loureiro@setuptecnologia.com.br'
senhaLogin = 'Racewin@1234'

# Caminho para o driver do Edge
driver_path = 'msedgedriver.exe'

# Configura as opções do Edge
edge_options = Options()
edge_options.add_argument('--log-level=3')  # Define o nível de log para 'fatal', suprimindo a maioria das mensagens de erro.
edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Exclui certos logs.

# Inicializa o navegador Edge
service = Service(driver_path)
edge_driver = webdriver.Edge(service=service, options=edge_options)
edge_driver.set_window_size(1300, 800)

# Abre o link desejado
url = f"https://app.acessorias.com/sysmain.php?m=4"
edge_driver.get(url)

# Lógica de login
try:
    # Espera o campo de e-mail aparecer
    email_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.NAME, 'mailAC'))
    )
    # Insere o e-mail no campo
    email_input.send_keys(emailLogin)
except:
    print("Erro ao inserir o e-mail no campo de login:")
    

try:
    # Espera o campo de senha aparecer
    senha_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.NAME, 'passAC'))
    )
    # Insere a senha no campo
    senha_input.send_keys(senhaLogin)
except:
    print("Erro ao inserir a senha no campo de senha:")
    

try:
    # Espera o botão de login aparecer e ser clicável
    login_button = WebDriverWait(edge_driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="site-corpo"]/section[1]/div/form/div[2]/button'))
    )
    # Clique no botão de login
    login_button.click()
except:
    print("Erro ao clicar no botão de login:")

time.sleep(2)

# Carrega o arquivo .xlsx
workbook = openpyxl.load_workbook('teste.xlsx')

# Seleciona a planilha
sheet = workbook.active  # Corrige aqui

# Função para normalizar o texto, substituindo 'º' por '°'
def normalize_text(text):
    # Remove acentos e caracteres especiais
    normalized = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII').lower()
    # Substitui caracteres específicos
    normalized = normalized.replace('º', '°')
    # Considera substituições adicionais, se necessário
    # Exemplo: convertendo caracteres acentuados
    normalized = normalized.replace('ã', 'a').replace('á', 'a').replace('à', 'a').replace('â', 'a')
    normalized = normalized.replace('é', 'e').replace('è', 'e').replace('ê', 'e')
    return normalized

for row in sheet.iter_rows(min_row=2, min_col=1, max_col=7):  # E aqui
    GrupoEmpresa = row[0].value          # Coluna 1: grupo de empresas
    RazaoSocial = row[1].value        # Coluna 3: Razão Social
    Cnpj = row[2].value               # Coluna 4: CNPJ/CPF/CEI/CAEPF
    Apelido = row[3].value            # Coluna 5: Insc. Estadual
    tagContabil = row[4].value             # Coluna 6: Insc. Estadual UF
    tagFiscal = row[5].value           # Coluna 7: Nome do Contato
    tagDP = row[6].value          # Coluna 8: E-mail do Contato
    
    url = 'https://app.acessorias.com/sysmain.php?m=4'
    edge_driver.get(url)
    
    try:
        # Executa o script JavaScript para chamar a função addEmp(this)
        edge_driver.execute_script("addEmp(this);")
    except Exception as e:
        print("Erro ao chamar a função addEmp(this):", e)
    
    try:
        while True:
            # Encontra o seletor <select> pelo nome
            selectGrupoEmpresa = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.NAME, 'field_EmpEgpID'))
            )

            # Cria uma instância de Select com o elemento encontrado
            select = Select(selectGrupoEmpresa)  
            grupo_normalized = normalize_text(GrupoEmpresa)
            option_found = False  # Flag para controlar se a opção foi encontrada
            
            # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável grupo empresa
            for option in select.options:
                option_normalized = normalize_text(option.text)
                if option_normalized == grupo_normalized:
                    select.select_by_visible_text(option.text)
                    option_found = True
                    break
            
            if option_found:
                break  # Sai do loop principal se a opção foi encontrada
            
            else:
                print(f"Opção '{grupo_normalized}' não encontrada, então vou tentar criar um por aqui")
                url = 'https://app.acessorias.com/sysmain.php?m=146&act=a'
                edge_driver.get(url)
                
                try:
                    # Espera o campo aparecer
                    NomeGrupoCriar = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, 'EgpNome'))
                    )
                    
                    # Insere o Apelido no campo
                    NomeGrupoCriar.send_keys(GrupoEmpresa)
                except Exception as e:
                    print(f"Erro ao inserir o nome do grupo: {e}")
                
                edge_driver.execute_script("check_form(this);")
                # Volta para a página original onde o select está localizado
                edge_driver.get('https://app.acessorias.com/sysmain.php?m=4')
                time.sleep(1)
                try:
                    # Executa o script JavaScript para chamar a função addEmp(this)
                    edge_driver.execute_script("addEmp(this);")
                except Exception as e:
                    print("Erro ao chamar a função addEmp(this):", e)
    except Exception as e:
        print(f"Erro: {e}")
    
    

    #inserir cnpj
    try:
        # Espera o campo aparecer
        cnpjInput = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'field_EmpCNPJ'))
        )
        
        # Insere o cnpj no campo
        cnpjInput.send_keys(Cnpj)
        cnpjInput.send_keys(Keys.TAB)
    except:
        print("Erro ao inserir o CNPJ na empresa")
        

    try:
        # Tenta localizar o botão "Buscar" com o ID "btCNPJ"
        buscarButton = WebDriverWait(edge_driver, 2).until(
            EC.element_to_be_clickable((By.ID, 'btCNPJ'))
        )
    except:
        try:
            # Se o botão "btCNPJ" não for encontrado, tenta localizar o botão com o ID "btCPF"
            buscarButton = WebDriverWait(edge_driver, 2).until(
                EC.element_to_be_clickable((By.ID, 'btCPF'))
            )
        except:
            print("Nenhum botão 'Buscar' encontrado")
            buscarButton = None

    # Se o botão foi encontrado, clique nele
    if buscarButton:
        buscarButton.click()

        try:
            # Aguarda até que o pop-up de erro apareça
            popUpErro = WebDriverWait(edge_driver, 2).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, '.swal2-popup.swal2-modal'))
            )
            print("Pop-up de erro encontrado!")
            # Aqui você pode adicionar código para interagir com o pop-up, se necessário

        except:
            print("Pop-up de erro não apareceu")
    else:
        print("O botão 'Buscar' não foi clicado, portanto, o pop-up não apareceu.")
        
    # visibilidade do elemento que contém a mensagem de erro
    try:
        popUpErro = WebDriverWait(edge_driver, 2).until(
            EC.visibility_of_element_located((By.XPATH, '//div[@aria-labelledby="swal2-title"]'))
        )

        # Se precisar "Continuar":
        if popUpErro:
            botao_continuar = popUpErro.find_element(By.XPATH, './/button[contains(@class, "swal2-confirm")]')
            botao_continuar.click()  # Clica no botão "Continuar"
        else:
            break
    except:
        print(" ")

    if RazaoSocial:
        try:
            # Espera o campo de IE da empresa aparecer
            razaoInput = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.NAME, 'field_EmpNome'))
            )
            
            # Insere o ID da empresa no campo
            razaoInput.clear()
            razaoInput.send_keys(RazaoSocial)

            # Espera o campo de IE da empresa aparecer
            fantasiaInput = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.NAME, 'field_EmpFantasia'))
            )
            
            # Insere o ID da empresa no campo
            fantasiaInput.clear()
            fantasiaInput.send_keys(RazaoSocial)

        except Exception as e:
            print("Erro ao colocar a razao social.")

    if Apelido:    
        #inserir Apelido
        try:
            # Espera o campo aparecer
            ApelidoEcont = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="EmpApelido"]'))
            )
            
            # Insere o Apelido no campo
            ApelidoEcont.send_keys(Apelido)
        except:
            print("Erro ao inserir o Apelido Econtínuo")
    

    # Espera até o ícone do grupo estar clicável
    tagIconElement = WebDriverWait(edge_driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'iDivTag'))
    )

    # Obtém a classe atual do ícone do grupo
    tagIconClass = tagIconElement.get_attribute("class")

    # Se a classe contém 'grey', clica no ícone para mudar para 'green'
    if 'grey' in tagIconClass:
        tagIconElement.click()

    if tagContabil:
        try:
            tagInput = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'tt-input'))
            )
            tagInput.send_keys(tagContabil)
            tagInput.send_keys(Keys.ENTER)
            time.sleep(0.5)
        except:
            print("erro ao incluir a tag" + tagContabil + 'na empresa de cnpj: ' + Cnpj)

    if tagFiscal:
        try:
            tagInput = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'tt-input'))
            )
            tagInput.send_keys(tagFiscal)
            tagInput.send_keys(Keys.ENTER)
            time.sleep(0.5)
        except:
            print("erro ao incluir a tag" + tagFiscal + 'na empresa de cnpj: ' + Cnpj)

    if tagDP:
        print(tagDP)
        try:
            tagInput = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'tt-input'))
            )
            tagInput.send_keys(tagDP)
            tagInput.send_keys(Keys.ENTER)
            time.sleep(0.5)
        except:
            print("erro ao incluir a tag " + tagDP + 'na empresa de cnpj: ' + Cnpj)

    try:
        # Espera o botão de salvar aparecer
        save_button = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button.btn.btn-sm.btn-success.col-xs-6.col-sm-6[onclick="check_form(this);"]'))
        )
        # Scroll até o botão usando JavaScript
        edge_driver.execute_script("arguments[0].scrollIntoView(true);", save_button)
        # Clica no botão
        time.sleep(0.5)
        # save_button.click()
        print('se eu quisesse salvar')
        time.sleep(2)
    except Exception as e:
        print("Não consegui cadastrar a empresa de CNPJ: " + Cnpj)
        
    
edge_driver.quit()
