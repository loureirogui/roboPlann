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
import unicodedata
import re


def atualizaObrigacao(emailLogin, senhaLogin):
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

                
        #Tabela com as obrigações    
        sheet = workbook['Obrigações']
                
                # Função para remover acentos e diacríticos
        def remove_acento(text):
            return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')

        # Função para normalizar o texto
        def normalize_text(text):
            if not isinstance(text, str):
                raise ValueError("A entrada deve ser uma string")
            
            # Remove acentos e caracteres diacríticos
            text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
            
            # Substitui o símbolo ordinal "º" e "°" por "o"
            text = text.replace('º', 'o').replace('°', 'o')
            
            # Remove todos os caracteres não alfanuméricos e não espaços, preservando espaços
            text = re.sub(r'[^\w\s]', '', text)  # Usando \w para preservar letras e números
            
            # Remove espaços adicionais e converte para minúsculas
            text = ' '.join(text.split()).lower()
            
            return text

        for row in sheet.iter_rows(min_row=4, min_col=1, max_col=20):
            
            NomeObrigacao = row[0].value
            Dpto = row[1].value
            Janeiro = row[2].value
            Fevereiro = row[3].value
            Marco = row[4].value
            Abril = row[5].value
            Maio = row[6].value
            Junho = row[7].value
            Julho = row[8].value
            Agosto = row[9].value
            Setembro = row[10].value
            Outubro = row[11].value
            Novembro = row[12].value
            Dezembro = row[13].value
            PrazoTec = row[14].value
            Dias = row[15].value
            comp = row[18].value
            Multa = row[19].value


            if NomeObrigacao:
                try:
                    url = f"https://app.acessorias.com/sysmain.php?m=20"
                    edge_driver.get(url)
                    time.sleep(2)
                    
                    try:
                        # Espera o campo de nome aparecer
                        searchObrigacao = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'search'))
                        )
                        # Insere o nome no campo
                        searchObrigacao.send_keys(NomeObrigacao)
                    except Exception:
                        print("Erro ao inserir o nome do usuário")


                    try:
                        # Espera o botão de filtro aparecer
                        filtrarButton = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.ID, 'btFilter'))
                        )
                        # Clica no botão de filtro
                        filtrarButton.click()
                    except Exception as e:
                        print("Erro ao clicar no botão de filtrar:", e)

                    try:
                        # Aguarde até que os elementos estejam presentes
                        divs = WebDriverWait(edge_driver, 10).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".col-xs-12.col-sm-12.dRow.aImage, .col-xs-12.col-sm-12.dOdd.aImage"))
                        )

                        # Itera sobre todas as divs encontradas
                        for div in divs:
                            try:
                                # Encontra o span com a classe 'blue' dentro da div atual
                                span = div.find_element(By.XPATH, '//*[@id="divList"]/div[2]/div[1]/div[1]/span[1]')
                                
                                # Verifica se o texto dentro do span corresponde ao NomeObrigacao
                                if span.text.strip() == NomeObrigacao:
                                    # Se encontrar a div correta, faça algo com ela
                                    
                                    
                                    # Exemplo de ação: Adicionar uma classe ao elemento (via JavaScript)
                                    span.click()
                                    
                                    def get_first_word(text):
                                        return text.split()[0]

                            

                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        select_element = WebDriverWait(edge_driver, 10).until(
                                        EC.visibility_of_element_located((By.NAME, 'ObrDptID')))
                                        
                                        
                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(select_element)
                                        
                                        # Obtém a primeira palavra do Dpto
                                        first_word_dpto = get_first_word(Dpto)
                                        
                                        # Itera através das opções para encontrar aquela cuja primeira palavra do texto corresponde à primeira palavra do Dpto
                                        for option in select.options:
                                            first_word_option = get_first_word(option.text)
                                            if first_word_option == first_word_dpto:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print("Opção não encontrada")
                                    except Exception as e:
                                        print(f"Erro: {e}")



                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaJaneiro = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD01'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaJaneiro)

                                        # Normaliza a variável Janeiro para comparação
                                        janeiro_normalized = normalize_text(Janeiro)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Janeiro
                                        found = False
                                        for option in select.options:
                                            # Extraí o texto da opção
                                            option_text = option.text.strip()  # Remove espaços em branco extras
                                            
                                            # Normaliza o texto da opção
                                            option_normalized = normalize_text(option_text)
                                            
                                            # Print para depuração
                                            
                                            
                                            if option_normalized == janeiro_normalized:
                                                select.select_by_visible_text(option.text)
                                                found = True
                                                break
                                        
                                        if not found:
                                            print(f"Opção '{Janeiro}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")
                                    #Entrega Fevereiro
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaFevereiro = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD02'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaFevereiro)

                                        # Define a variável Fevereiro (substitua 'Fevereiro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Fevereiro para comparação
                                        Fevereiro_normalized = normalize_text(Fevereiro)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Fevereiro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Fevereiro_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Fevereiro}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Entrega Marco
                                    
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaMarco = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD03'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaMarco)

                                        # Define a variável Março (substitua 'Março' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Março para comparação
                                        Marco_normalized = normalize_text(Marco)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Março
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Marco_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Marco}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Entrega Abril
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaAbril = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD04'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaAbril)

                                        # Define a variável Abril (substitua 'Abril' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Abril para comparação
                                        Abril_normalized = normalize_text(Abril)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Abril
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Abril_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Abril}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")
                                    
                                    #Entrega Maio

                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaMaio = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD05'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaMaio)

                                        # Define a variável Maio (substitua 'Maio' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Maio para comparação
                                        Maio_normalized = normalize_text(Maio)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Maio
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Maio_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Maio}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")
                                    
                                    #Entrega Junho
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaJunho = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD06'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaJunho)

                                        # Define a variável Junho (substitua 'Junho' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Junho para comparação
                                        Junho_normalized = normalize_text(Junho)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Junho
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Junho_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Junho}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Julho
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaJulho = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD07'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaJulho)

                                        # Define a variável Julho (substitua 'Julho' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Julho para comparação
                                        Julho_normalized = normalize_text(Julho)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Julho
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Julho_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Julho}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")
                                    
                                #Agosto
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaAgosto = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD08'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaAgosto)

                                        # Define a variável Agosto (substitua 'Agosto' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Agosto para comparação
                                        Agosto_normalized = normalize_text(Agosto)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Agosto
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Agosto_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Agosto}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Setembro 
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaSetembro = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD09'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaSetembro)

                                        # Define a variável Setembro (substitua 'Setembro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Setembro para comparação
                                        Setembro_normalized = normalize_text(Setembro)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Setembro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Setembro_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Setembro}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")


                                    #Outubro
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaOutubro = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD10'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaOutubro)

                                        # Define a variável Outubro (substitua 'Outubro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Outubro para comparação
                                        Outubro_normalized = normalize_text(Outubro)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Outubro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Outubro_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Outubro}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Novembro
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaNovembro = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD11'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaNovembro)

                                        # Define a variável Novembro (substitua 'Novembro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Novembro para comparação
                                        Novembro_normalized = normalize_text(Novembro)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Novembro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Novembro_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Novembro}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Dezembro
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        entregaDezembro = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrD12'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(entregaDezembro)

                                        # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Dezembro para comparação
                                        Dezembro_normalized = normalize_text(Dezembro)

                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            if option_normalized == Dezembro_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Dezembro}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}"),
                                
                                    #Lembrar responsável dias antes
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        PrazoTecnico = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrDAntes'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(PrazoTecnico)

                                        # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Dezembro para comparação
                                        Prazotec_normalized = normalize_text(PrazoTec)
                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            
                                            if option_normalized == Prazotec_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{PrazoTec}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Tipo de dias antes
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        tipoDiasAntes = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrDAntesTipo'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(tipoDiasAntes)

                                        # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Dezembro para comparação
                                        tipoDiaOpc_normalized = normalize_text(Dias)
                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            
                                            if option_normalized == tipoDiaOpc_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{Dias}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Competencia referente?
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        compReferen = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrCompetencia'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(compReferen)

                                        # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                        
                                        # Normaliza a variável Dezembro para comparação
                                        compReferen_normalized = normalize_text(comp)
                                        # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                        for option in select.options:
                                            option_normalized = normalize_text(option.text)
                                            
                                            if option_normalized == compReferen_normalized:
                                                select.select_by_visible_text(option.text)
                                                break
                                        else:
                                            print(f"Opção '{comp}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")


                                    #Passivel de multa?
                                    try:
                                        # Encontra o seletor <select> pelo nome
                                        multaOpt = WebDriverWait(edge_driver, 10).until(
                                            EC.visibility_of_element_located((By.NAME, 'ObrMulta'))
                                        )

                                        # Cria uma instância de Select com o elemento encontrado
                                        select = Select(multaOpt)

                                        # Normaliza a variável Multa para comparação
                                        multa_normalized = normalize_text(Multa)

                                        # Flag para verificar se a opção foi encontrada
                                        option_found = False

                                        # Itera através das opções para encontrar a opção desejada
                                        for option in select.options:
                                            option_text_normalized = normalize_text(option.text)
                                            
                                            if option_text_normalized == multa_normalized:
                                                select.select_by_visible_text(option.text)
                                                option_found = True
                                                break

                                        if not option_found:
                                            print(f"Opção '{Multa}' não encontrada")

                                    except Exception as e:
                                        print(f"Erro: {e}")

                                    #Salvar alteração
                                    try:
                                        # Espera o botão de salvar aparecer
                                        save_button = WebDriverWait(edge_driver, 10).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-sm.btn-success.col-xs-4.col-sm-4'))
                                        )
                                    # Clique no botão de salvar
                                        save_button.click()
                                    except Exception as e:
                                        print(f"Erro ao clicar no botão de salvar: {e}")



                                break
                            except Exception as inner_e:
                                


                                print("Erro ao processar uma das divs:", inner_e)
                                continue

                    except Exception as e:
                        try:
                            # Espera o botão "Nova obrigação" aparecer e ser clicável
                            newObr_button = WebDriverWait(edge_driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-sm.btn-primary.col-xs-12.col-sm-2'))
                            )
                            # Clique no botão "Nova obrigação"
                            newObr_button.click()
                            
                            def get_first_word(text):
                                        return text.split()[0]

                            #Nome da obrigação
                            try:
                                # Espera o campo de e-mail aparecer
                                nomeObr_input = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrNome'))
                                )
                                #Limpa o nome só por garantia
                                nomeObr_input.clear()
                                # Insere o e-mail no campo
                                nomeObr_input.send_keys(NomeObrigacao)
                            except Exception:
                                print("Erro ao inserir o nome da obrigação")
                            
                            
                            try:
                                # Encontra o seletor <select> pelo nome
                                select_element = WebDriverWait(edge_driver, 10).until(
                                EC.visibility_of_element_located((By.NAME, 'ObrDptID')))
                                
                                
                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(select_element)
                                
                                # Obtém a primeira palavra do Dpto
                                first_word_dpto = get_first_word(Dpto)
                                
                                # Itera através das opções para encontrar aquela cuja primeira palavra do texto corresponde à primeira palavra do Dpto
                                for option in select.options:
                                    first_word_option = get_first_word(option.text)
                                    if first_word_option == first_word_dpto:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print("Opção não encontrada")
                            except Exception as e:
                                print(f"Erro: {e}")



                            #Entrega Janeiro
                            # Função para remover acentos e diacríticos, incluindo o til
                            def remove_acento(text):
                                return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')

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
                            
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaJaneiro = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD01'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaJaneiro)

                                # Define a variável Janeiro (substitua 'Janeiro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Janeiro para comparação
                                Janeiro_normalized = normalize_text(Janeiro)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Janeiro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    
                                    if option_normalized == Janeiro_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Janeiro}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Entrega Fevereir
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaFevereiro = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD02'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaFevereiro)

                                # Define a variável Fevereiro (substitua 'Fevereiro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Fevereiro para comparação
                                Fevereiro_normalized = normalize_text(Fevereiro)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Fevereiro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Fevereiro_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Fevereiro}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Entrega Marco
                            
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaMarco = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD03'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaMarco)

                                # Define a variável Março (substitua 'Março' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Março para comparação
                                Marco_normalized = normalize_text(Marco)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Março
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Marco_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Marco}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Entrega Abril
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaAbril = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD04'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaAbril)

                                # Define a variável Abril (substitua 'Abril' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Abril para comparação
                                Abril_normalized = normalize_text(Abril)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Abril
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Abril_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Abril}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")
                            
                            #Entrega Maio

                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaMaio = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD05'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaMaio)

                                # Define a variável Maio (substitua 'Maio' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Maio para comparação
                                Maio_normalized = normalize_text(Maio)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Maio
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Maio_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Maio}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")
                            
                            #Entrega Junho
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaJunho = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD06'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaJunho)

                                # Define a variável Junho (substitua 'Junho' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Junho para comparação
                                Junho_normalized = normalize_text(Junho)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Junho
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Junho_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Junho}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Julho
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaJulho = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD07'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaJulho)

                                # Define a variável Julho (substitua 'Julho' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Julho para comparação
                                Julho_normalized = normalize_text(Julho)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Julho
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Julho_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Julho}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")
                            
                            #Agosto
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaAgosto = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD08'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaAgosto)

                                # Define a variável Agosto (substitua 'Agosto' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Agosto para comparação
                                Agosto_normalized = normalize_text(Agosto)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Agosto
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Agosto_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Agosto}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Setembro 
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaSetembro = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD09'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaSetembro)

                                # Define a variável Setembro (substitua 'Setembro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Setembro para comparação
                                Setembro_normalized = normalize_text(Setembro)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Setembro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Setembro_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Setembro}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")


                            #Outubro
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaOutubro = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD10'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaOutubro)

                                # Define a variável Outubro (substitua 'Outubro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Outubro para comparação
                                Outubro_normalized = normalize_text(Outubro)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Outubro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Outubro_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Outubro}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Novembro
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaNovembro = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD11'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaNovembro)

                                # Define a variável Novembro (substitua 'Novembro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Novembro para comparação
                                Novembro_normalized = normalize_text(Novembro)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Novembro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Novembro_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Novembro}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Dezembro
                            try:
                                # Encontra o seletor <select> pelo nome
                                entregaDezembro = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrD12'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(entregaDezembro)

                                # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Dezembro para comparação
                                Dezembro_normalized = normalize_text(Dezembro)

                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    if option_normalized == Dezembro_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Dezembro}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}"),
                        
                            #Lembrar responsável dias antes
                            try:
                                # Encontra o seletor <select> pelo nome
                                PrazoTecnico = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrDAntes'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(PrazoTecnico)

                                # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Dezembro para comparação
                                Prazotec_normalized = normalize_text(PrazoTec)
                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    
                                    if option_normalized == Prazotec_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{PrazoTec}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Tipo de dias antes
                            try:
                                # Encontra o seletor <select> pelo nome
                                tipoDiasAntes = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrDAntesTipo'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(tipoDiasAntes)

                                # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Dezembro para comparação
                                tipoDiaOpc_normalized = normalize_text(Dias)
                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    
                                    if option_normalized == tipoDiaOpc_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{Dias}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Competencia referente?
                            try:
                                # Encontra o seletor <select> pelo nome
                                compReferen = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrCompetencia'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(compReferen)

                                # Define a variável Dezembro (substitua 'Dezembro' pelo valor real que você deseja usar)
                                
                                # Normaliza a variável Dezembro para comparação
                                compReferen_normalized = normalize_text(comp)
                                # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Dezembro
                                for option in select.options:
                                    option_normalized = normalize_text(option.text)
                                    
                                    if option_normalized == compReferen_normalized:
                                        select.select_by_visible_text(option.text)
                                        break
                                else:
                                    print(f"Opção '{comp}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")


                            #Passivel de multa?
                            try:
                                # Encontra o seletor <select> pelo nome
                                multaOpt = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrMulta'))
                                )

                                # Cria uma instância de Select com o elemento encontrado
                                select = Select(multaOpt)

                                # Normaliza a variável Multa para comparação
                                multa_normalized = normalize_text(Multa)

                                # Flag para verificar se a opção foi encontrada
                                option_found = False

                                # Itera através das opções para encontrar a opção desejada
                                for option in select.options:
                                    option_text_normalized = normalize_text(option.text)
                                    
                                    if option_text_normalized == multa_normalized:
                                        select.select_by_visible_text(option.text)
                                        option_found = True
                                        break

                                if not option_found:
                                    print(f"Opção '{Multa}' não encontrada")

                            except Exception as e:
                                print(f"Erro: {e}")

                            #Salvar alteração
                            try:
                                # Espera o botão de salvar aparecer
                                save_button = WebDriverWait(edge_driver, 10).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-sm.btn-success.col-xs-4.col-sm-4'))
                                )
                                # Clique no botão de salvar
                                save_button.click()
                                time.sleep(2)
                            except Exception as e:
                                print(f"Erro ao clicar no botão de salvar: {e}")
                        except Exception as e:
                            print(f"Erro ao clicar no botão de adicionar obrigação: {e}")
                
                except Exception as e:
                    print('Erro ao abrir a URL:', e)
    except Exception as e:
            print('Erro ao abrir a URL:', e)