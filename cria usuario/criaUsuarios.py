import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from docx import Document
import openpyxl


def criaUsuario(emailLogin, senhaLogin):
    try:
        
        # Caminho para o driver do Edge
        driver_path = 'msedgedriver.exe'

        # Configura as opções do Edge
        edge_options = Options()
        edge_options.add_argument('--log-level=3')  # Isso define o nível de log para 'fatal', suprimindo a maioria das mensagens de erro.
        edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Exclui certos logs.

        # Inicialize o EdgeDriver com as opções configuradas


        # Inicializa o navegador Edge
        service = Service(driver_path)
        edge_driver = webdriver.Edge(service=service, options=edge_options)

        # Abre o link desejado
        url = f"https://app.acessorias.com/sysmain.php?m=105&act=e&i=365&uP=14&o=EmpNome,EmpID|Asc"
        edge_driver.get(url)

        # Carrega o arquivo .xlsx
        workbook = openpyxl.load_workbook('uploaded_empresas.xlsx')

        # Seleciona a planilha ativa (a primeira planilha aberta por padrão)
        sheet = workbook['Colaboradores']

        # Seleciona as colunas específicas
        colunaNome = 'A'
        colunaEmail = 'B'

        # Lógica de login
        try:
            # Espera o campo de e-mail aparecer
            email_input = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.NAME, 'mailAC'))
            )
            # Insere o e-mail no campo
            email_input.send_keys(emailLogin)
        except Exception:
            print("Erro ao inserir o e-mail no campo de login:")

        try:
            # Espera o campo de senha aparecer
            senha_input = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.NAME, 'passAC'))
            )
            # Insere a senha no campo
            senha_input.send_keys(senhaLogin)
        except Exception:
            print("Erro ao inserir a senha no campo de senha:")

        # Espera o botão de login aparecer
        try:
            login_button = WebDriverWait(edge_driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.button.rounded.large.expanded.primary-degrade.btn-enviar'))
            )
            # Clique no botão de login
            login_button.click()
        except Exception:
            print("Erro ao clicar no botão de login:")

        # Espera 2 segundos após clicar no botão de login
        time.sleep(2)

        for row in sheet.iter_rows(min_row=3, min_col=1, max_col=2):
            nomePlanilha = row[0].value
            emailPlanilha = row[1].value
            if nomePlanilha and emailPlanilha:
                try:
                    url = f"https://app.acessorias.com/sysmain.php?m=16&act=a"
                    edge_driver.get(url)
                    time.sleep(2)

                    try:
                        # Espera o campo de nome aparecer
                        nome_input = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'LogNome'))
                        )
                        # Insere o nome no campo
                        nome_input.send_keys(nomePlanilha)
                    except Exception:
                        print("Erro ao inserir o nome do usuário")

                    try:
                        # Espera o campo de email aparecer
                        email_usuario = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'LogEmail'))
                        )
                        # Insere o email no campo
                        email_usuario.send_keys(emailPlanilha)
                    except Exception:
                        print("Erro ao inserir o email do usuário")

                    try:
                        # Espera o campo de email aparecer
                        tipo_usuario = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'LogTipo'))
                        )

                        # Seleciona o campo <select> pelo id
                        select_tipo_usuario = Select(tipo_usuario)

                        # Seleciona a opção "Contador sócio" pelo valor
                        select_tipo_usuario.select_by_value('P')

                    except Exception as e:
                        print("Erro ao selecionar o tipo de usuário:", e)
                        
                    try:
                        # Espera o botão de salvar aparecer
                        save_button = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button.btn.btn-sm.btn-success.col-xs-6.col-sm-6[onclick="check_form(this);"]'))
                        )
                        # Scroll até o botão usando JavaScript
                        edge_driver.execute_script("arguments[0].scrollIntoView(true);", save_button)
                        # Clica no botão
                        time.sleep(2)
                        save_button.click()
                        time.sleep(2)
                    except Exception as e:
                        print("Erro ao clicar no botão de salvar:", e)
                except:
                    print('Erro ao abrir a URL')
    except:
        print('Erro ao abrir a URL')
                