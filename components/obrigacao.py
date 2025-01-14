import traceback
import time
import difflib
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
import os
from webdriver_manager.microsoft import EdgeChromiumDriverManager

def registrar_erro(nomeCliente, mensagem):
    # Cria o nome do diretório com base no nome do cliente
    pasta_cliente = f"./{nomeCliente}"  # ./ indica que a pasta será criada no diretório atual
    
    # Verifica se a pasta já existe, se não, cria
    if not os.path.exists(pasta_cliente):
        os.makedirs(pasta_cliente)
    
    # Define o caminho completo do arquivo de erro dentro da pasta do cliente
    erro_log = os.path.join(pasta_cliente, f"erros_{nomeCliente}_obriga.txt")
    
    # Abre o arquivo no modo de apêndice e escreve a mensagem de erro
    with open(erro_log, "a", encoding="utf-8") as file:
        file.write(mensagem + "\n")

def atualizaObrigacao(nomeCliente, emailLogin, senhaLogin, nomePlanilha):
    # Caminho para o driver do Edge
    driver_path = 'msedgedriver.exe'

    # Configura as opções do Edge
    edge_options = Options()
    edge_options.add_argument('--log-level=3')  # Isso define o nível de log para 'fatal', suprimindo a maioria das mensagens de erro.
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Exclui certos logs.

    # Inicialize o EdgeDriver com as opções configuradas


    # Inicializa o navegador Edge
    #service = Service(driver_path)
    edge_driver = webdriver.Edge(    service=Service(EdgeChromiumDriverManager().install()), options=edge_options)

    # Abre o link desejado
    url = f"https://app.acessorias.com/sysmain.php?m=105&act=e&i=365&uP=14&o=EmpNome,EmpID|Asc"
    edge_driver.get(url)

    # Carrega o arquivo .xlsx
    workbook = openpyxl.load_workbook(nomePlanilha)

    # Seleciona a planilha ativa (a primeira planilha aberta por padrão)
    sheet = workbook['Obrigações']

    # Seleciona as colunas específicas
    colunaNomeObrigacao = 'A'
    colunaDpto = 'B'
    colunaJaneiro = 'C'
    colunaFevereiro = 'D'
    colunaMarco = 'E'
    colunaAbril = 'F'
    colunaMaio = 'G'
    colunaJunho = 'H'
    colunaJulho = 'I'
    colunaAgosto = 'J'
    colunaSetembro = 'K'
    colunaOutubro = 'L'
    colunaNovembro = 'M'
    colunaDezembro = 'N'
    colunaPrazoTec = 'O'
    tipoDias = 'P'
    competencia = 'S'
    sofreMulta = 'T'

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
        text = text.replace('º', '').replace('°', '')
        text = text.replace('ã', 'a').replace('ú', 'u')
        text = text.replace('dias', 'dia')
        # Remove todos os caracteres nao alfanuméricos e nao espaços, preservando espaços
            
        # Remove espaços adicionais e converte para minúsculas
        text = ' '.join(text.split()).lower()
        
        return text
        

    for row in sheet.iter_rows(min_row=4, min_col=0, max_col=21):
        
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


        # Aqui você continua o seu código para processar a linha...
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
                    print("Erro ao inserir o nome da obrigacao")
                    registrar_erro(nomeCliente,"Erro ao inserir o nome da obrigacao;" + NomeObrigacao)

                try:
                    # Espera o botão de filtro aparecer
                    filtrarButton = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.ID, 'btFilter'))
                    )
                    # Clica no botão de filtro
                    filtrarButton.click()
                except Exception:
                    print("Erro ao clicar no botao de filtrar:")
                    
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
                                span.click()
                        
                                def get_first_word(text):
                                    return text.split()[0]

                        

                                try:
                                    # Encontra o seletor <select> pelo nome
                                    select_element = WebDriverWait(edge_driver, 10).until(
                                    EC.visibility_of_element_located((By.NAME, 'ObrDptID')))
            
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
                                        print("Opcao;" + Dpto + ";nao encontrada")
                                        registrar_erro(nomeCliente,"Opcao;" + Dpto + ";nao encontrada")
                                except Exception:
                                    print(f"Erro:")



                                try:
                                    # Encontra o seletor <select> pelo nome
                                    entregaJaneiro = WebDriverWait(edge_driver, 10).until(
                                        EC.visibility_of_element_located((By.NAME, 'ObrD01'))
                                    )

                                    select = Select(entregaJaneiro)

                                    # Normaliza a variável Janeiro para comparação
                                    janeiro_normalized = normalize_text(Janeiro)

                                    # Itera através das opções para encontrar aquela cujo texto é igual ao valor da variável Janeiro
                                    found = False
                                    for option in select.options:
                                        # Extraí o texto da Opcao
                                        option_text = option.text.strip()  # Remove espaços em branco extras
                                        
                                        # Normaliza o texto da Opcao
                                        option_normalized = normalize_text(option_text)
                                        
                                        # Print para depuração
                                        
                                        
                                        if option_normalized == janeiro_normalized:
                                            select.select_by_visible_text(option.text)
                                            found = True
                                            break
                                    
                                    if not found:
                                        print(f"Opcao nao encontrada no campo ;" + janeiro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + janeiro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                except Exception as e:
                                    print(f"Erro:")
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
                                        print(f"Opcao nao encontrada no campo ;" + Fevereiro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Fevereiro_normalized + ";Para a obrigacao;" + NomeObrigacao)


                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao nao encontrada no campo ;" + Marco_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Marco_normalized + ";Para a obrigacao;" + NomeObrigacao)


                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao nao encontrada no campo ;" + Abril_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Abril_normalized + ";Para a obrigacao;" + NomeObrigacao)

                                except Exception as e:
                                    print(f"Erro:")
                                
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
                                        print(f"Opcao nao encontrada no campo ;" + Maio_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Maio_normalized + ";Para a obrigacao;" + NomeObrigacao)

                                except Exception as e:
                                    print(f"Erro:")
                                
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
                                        print(f"Opcao nao encontrada no campo ;" + Junho_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Junho_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao nao encontrada no campo ;" + Julho_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Julho_normalized + ";Para a obrigacao;" + NomeObrigacao)

                                except Exception as e:
                                    print(f"Erro:")
                                
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
                                        print(f"Opcao nao encontrada no campo Janeiro ;" + Agosto_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo Janeiro ;" + Agosto_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao '{Setembro}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Setembro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"Erro:")


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
                                        print(f"Opcao '{Outubro}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Outubro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao '{Novembro}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Novembro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    

                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao '{Dezembro}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Dezembro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"Erro:"),
                            
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
                                        print(f"Opcao '{PrazoTec}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Prazotec_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao '{Dias}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + tipoDiaOpc_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"Erro:")

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
                                        print(f"Opcao '{comp}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + compReferen_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"Erro:")


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

                                    # Flag para verificar se a Opcao foi encontrada
                                    option_found = False

                                    # Itera através das opções para encontrar a Opcao desejada
                                    for option in select.options:
                                        option_text_normalized = normalize_text(option.text)
                                        
                                        if option_text_normalized == multa_normalized:
                                            select.select_by_visible_text(option.text)
                                            option_found = True
                                            break

                                    if not option_found:
                                        print(f"Opcao '{Multa}' nao encontrada")
                                        registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + multa_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"Erro:")

                                #Salvar alteração
                                try:
                                    edge_driver.execute_script("check_form(this);")
                                    print(f"Obrigação salva com sucesso;" + NomeObrigacao)
                                    registrar_erro(f"Obrigação salva com sucesso;" + NomeObrigacao)
                                    
                                except Exception as e:
                                    print(f"")
                                    

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
                            print("Erro ao inserir o nome da obrigacao")
                            registrar_erro(nomeCliente,"Erro ao inserir o nome da obrigacao;" + NomeObrigacao)

                        
                        
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
                                print("Opcao;" + Dpto + ";nao encontrada")
                                registrar_erro(nomeCliente,"Opcao;" + Dpto + ";nao encontrada")
                        except Exception as e:
                            print(f"Erro:")



                        
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
                                print(f"Opcao nao encontrada no campo ;" + janeiro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + janeiro_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao nao encontrada no campo ;" + Fevereiro_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Fevereiro_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao nao encontrada no campo ;" + Marco_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Marco_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao nao encontrada no campo ;" + Abril_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Abril_normalized + ";Para a obrigacao;" + NomeObrigacao)

                        except Exception as e:
                            print(f"Erro:")
                        
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
                                print(f"Opcao nao encontrada no campo ;" + Maio_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Maio_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")
                        
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
                                print(f"Opcao nao encontrada no campo ;" + Junho_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Junho_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao nao encontrada no campo ;" + Julho_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Julho_normalized + ";Para a obrigacao;" + NomeObrigacao)



                        except Exception as e:
                            print(f"Erro:")
                        
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
                                print(f"Opcao nao encontrada no campo Janeiro ;" + Agosto_normalized + ";Para a obrigacao;" + NomeObrigacao)
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo Janeiro ;" + Agosto_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao '{Setembro}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Setembro_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")


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
                                print(f"Opcao '{Outubro}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Outubro_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao '{Novembro}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Novembro_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao '{Dezembro}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Dezembro_normalized + ";Para a obrigacao;" + NomeObrigacao)

                        except Exception as e:
                            print(f"Erro:"),
                    
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
                                print(f"Opcao '{PrazoTec}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + Prazotec_normalized + ";Para a obrigacao;" + NomeObrigacao)


                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao '{Dias}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + tipoDiaOpc_normalized + ";Para a obrigacao;" + NomeObrigacao)

                        except Exception as e:
                            print(f"Erro:")

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
                                print(f"Opcao '{comp}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + compReferen_normalized + ";Para a obrigacao;" + NomeObrigacao)

                        except Exception as e:
                            print(f"Erro:")


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

                            # Flag para verificar se a Opcao foi encontrada
                            option_found = False

                            # Itera através das opções para encontrar a Opcao desejada
                            for option in select.options:
                                option_text_normalized = normalize_text(option.text)
                                
                                if option_text_normalized == multa_normalized:
                                    select.select_by_visible_text(option.text)
                                    option_found = True
                                    break

                            if not option_found:
                                print(f"Opcao '{Multa}' nao encontrada")
                                registrar_erro(nomeCliente,f"Opcao nao encontrada no campo ;" + multa_normalized + ";Para a obrigacao;" + NomeObrigacao)

                        except Exception as e:
                            print(f"Erro:")

                        #Salvar alteração
                        try:
                            edge_driver.execute_script("check_form(this);")
                            print(f"Obrigação salva com sucesso;" + NomeObrigacao)
                            registrar_erro(f"Obrigação salva com sucesso;" + NomeObrigacao)
                        except Exception as e:
                            print(f"Erro ao clicar no botão de salvar:")
                    except Exception as e:
                        print("")
                        
            
            except Exception as e:
                print('Erro ao abrir a URL:', e)

