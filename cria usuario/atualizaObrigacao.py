import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import re
import unicodedata

# Definindo o arquivo de log de erros
erro_log = "erros.txt"

def registrar_erro(mensagem):
    with open(erro_log, "a") as file:
        file.write(mensagem + "\n")

emailLogin = 'suporte@planning.com.br'
senhaLogin = 'Suporte321'
driver_path = 'msedgedriver.exe'

# Configura as opções do Edge
edge_options = Options()
edge_options.add_argument('--log-level=3')
edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Inicializa o navegador Edge
service = Service(driver_path)
edge_driver = webdriver.Edge(service=service, options=edge_options)
edge_driver.set_window_size(1300, 800)

# Abre o link desejado
url = f"https://app.acessorias.com/sysmain.php?m=105&act=e&i=365&uP=14&o=EmpNome,EmpID|Asc"
edge_driver.get(url)

# Lógica de login
try:
    email_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.NAME, 'mailAC'))
    )
    email_input.send_keys(emailLogin)
except Exception:
    registrar_erro("Erro ao inserir o e-mail no campo de login:")

try:
    senha_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.NAME, 'passAC'))
    )
    senha_input.send_keys(senhaLogin)
except Exception:
    registrar_erro("Erro ao inserir a senha no campo de senha:")

try:
    login_button = WebDriverWait(edge_driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.button.rounded.large.expanded.primary-degrade.btn-enviar'))
    )
    login_button.click()
except Exception:
    registrar_erro("Erro ao clicar no botão de login:")

time.sleep(2)

# Carregando o arquivo Excel
workbook = openpyxl.load_workbook('tarefas.xlsx')
sheet = workbook.active

# Inicializando o WebDriver
url = f"https://app.acessorias.com/sysmain.php?m=20&act=D"
edge_driver.get(url)

# Variável para armazenar a última obrigação processada
ultima_obrigacao = None

# Iterando sobre as linhas da planilha agrupando por obrigação
for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
    empresa = row[0].value
    cnpj = row[1].value
    obrigacao = row[2].value
    
    def normalize_text(text):
        """ Remove acentos e normaliza o texto para comparação """
        text = text.lower()  # Converte para minúsculas
        text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')  # Remove acentos
        text = re.sub(r'\s+', ' ', text).strip()  # Remove espaços extras
        return text

    # Comparando a obrigação atual com a última processada
    if obrigacao == ultima_obrigacao:
        #Colocar a primeira empresa aqui pois se depender do ok vai pular a primeira empresa
        try:
            # Encontra o campo de pesquisa dentro do elemento <span>
            search_input = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'select2-search__field'))
            )
            search_input.clear()
            search_input.send_keys(empresa)  # Insere o nome da empresa

            # Aguarda até que a lista de resultados esteja visível
            result_elements = WebDriverWait(edge_driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.select2-results__option'))
            )
            
            empresa_selecionada = None

            normalized_empresa = normalize_text(empresa)

            for item in result_elements:
                try:
                    item_text = item.text.strip()
                    aria_selected = item.get_attribute('aria-selected') or 'true'  # Tratamento para casos sem o atributo
                    
                    normalized_item_text = normalize_text(item_text)

                    if normalized_empresa in normalized_item_text and aria_selected == 'false':
                        empresa_selecionada = item
                        break
                except:
                    registrar_erro(f"Erro ao processar item de resultado para a empresa '{empresa}'.")

            if empresa_selecionada:
                empresa_selecionada.click()
                registrar_erro(f"Empresa; '{empresa}'; de CNPJ; '{cnpj}'.")
            else:
                registrar_erro(f"Nao foi possivel adicionar obrigacao '{obrigacao}' na empresa '{empresa}'; Provavelmente essa obrigacao j esteja alocada na empresa, mas vale a pena conferir")

        except Exception as e:
            registrar_erro(f"Erro ao selecionar a empresa '{empresa}' para a obrigação '{obrigacao}': {e}")
    else:
        registrar_erro(f"Nova obrigação encontrada '{obrigacao}'; Vou começar a alocar ela nas empresas")
        url = f"https://app.acessorias.com/sysmain.php?m=20&act=D"
        edge_driver.get(url)
        
        # Alocar o nome da nova obrigação
        try:
            # Espera o campo de obrigação aparecer
            obrigacaoSelect = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.NAME, 'AvuObrID'))
            )

            # Seleciona o campo <select> pelo id
            select_tipo_usuario = Select(obrigacaoSelect)

            # Seleciona a opção pelo valor
            select_tipo_usuario.select_by_visible_text(obrigacao)
        except Exception as e:
            registrar_erro(f"Erro ao selecionar o nome da obrigação '{obrigacao}': {e}")
            
        try:
            # Encontra o campo de pesquisa dentro do elemento <span>
            search_input = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'select2-search__field'))
            )
            
            search_input.send_keys(empresa)  # Insere o nome da empresa

            # Aguarda até que a lista de resultados esteja visível e selecione a empresa correta
            result_xpath = f"//li[contains(text(), '{empresa}')]"
            empresa_selecionada = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, result_xpath))
            )
            empresa_selecionada.click()

        except Exception as e:
            registrar_erro(f"Erro ao selecionar a empresa '{empresa}': {e}")
            
    # Atualizando a última obrigação processada
    ultima_obrigacao = obrigacao

# Finaliza o navegador
edge_driver.quit()
