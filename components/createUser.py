import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
import os
from webdriver_manager.microsoft import EdgeChromiumDriverManager

def log_error(client_name, message):
    """
    Registra logs de erros em um arquivo de texto dentro de uma pasta específica do cliente.
    Permite análise posterior dos problemas ocorridos durante a execução.
    """
    client_folder = f"./{client_name}"
    if not os.path.exists(client_folder):
        os.makedirs(client_folder)
    
    error_log_path = os.path.join(client_folder, f"errors_{client_name}_users.txt")
    with open(error_log_path, "a", encoding="utf-8") as file:
        file.write(message + "\n")

def create_users(client_name, login_email, login_password, spreadsheet_name):

    try:
        # Configurações do navegador Edge
        edge_options = Options()
        edge_options.add_argument('--log-level=3')  # Define o nível de log para 'fatal', reduzindo mensagens não essenciais
        edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Exclui logs internos do navegador
        # Nota: Essas configurações foram aplicadas para minimizar o ruído no console durante a execução.
        # Caso seja necessário depurar, considere ajustar '--log-level' para um valor menos restritivo como '1' (info).
        browser = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=edge_options)

        # Abre o link inicial
        initial_url = "https://app.acessorias.com/sysmain.php?m=105&act=e&i=365&uP=14&o=EmpNome,EmpID|Asc"
        browser.get(initial_url)

        # Carrega a planilha .xlsx
        workbook = openpyxl.load_workbook(spreadsheet_name)
        sheet = workbook['Colaboradores']

        # Lógica de login na aplicação
        try:
            email_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.NAME, 'mailAC'))
            )
            email_field.send_keys(login_email)
        except Exception:
            print("Erro ao inserir o e-mail no campo de login.")
            log_error(client_name, "Erro ao inserir o e-mail no campo de login.")

        try:
            password_field = WebDriverWait(browser, 10).until(
                EC.visibility_of_element_located((By.NAME, 'passAC'))
            )
            password_field.send_keys(login_password)
        except Exception:
            print("Erro ao inserir a senha no campo de senha.")
            log_error(client_name, "Erro ao inserir a senha no campo de senha.")

        try:
            login_button = WebDriverWait(browser, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.button.rounded.large.expanded.primary-degrade.btn-enviar'))
            )
            login_button.click()
        except Exception:
            print("Erro ao clicar no botão de login.")
            log_error(client_name, "Erro ao clicar no botão de login.")

        time.sleep(2)  # Aguarda o carregamento da página após o login

        for row in sheet.iter_rows(min_row=3, min_col=1, max_col=2):
            user_name = row[0].value
            user_email = row[1].value

            if user_name and user_email:
                try:
                    user_creation_url = "https://app.acessorias.com/sysmain.php?m=16&act=a"
                    browser.get(user_creation_url)
                    time.sleep(0.5)

                    try:
                        name_field = WebDriverWait(browser, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'LogNome'))
                        )
                        name_field.send_keys(user_name)
                    except Exception:
                        print(f"Erro ao inserir o nome do usuário: {user_name}")
                        log_error(client_name, f"Erro ao inserir o nome do usuário: {user_name}")

                    try:
                        email_field = WebDriverWait(browser, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'LogEmail'))
                        )
                        email_field.send_keys(user_email)
                    except Exception:
                        print(f"Erro ao inserir o email do usuário: {user_email}")
                        log_error(client_name, f"Erro ao inserir o email do usuário: {user_email}")

                    try:
                        user_type_field = WebDriverWait(browser, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'LogTipo'))
                        )
                        select_user_type = Select(user_type_field)
                        select_user_type.select_by_value('P')  # Seleciona "Contador sócio". Por padrão, no inicio da implantação 
                        #é bom que todos os usuários tenham acesso elevado para ajudar no processo.
                        time.sleep(0.5)
                        browser.execute_script("check_form(this)")

                        print(f"Usuário criado com sucesso: {user_name}, {user_email}")
                        log_error(client_name, f"Usuário criado com sucesso: {user_name}, {user_email}")

                    except Exception as e:
                        print(f"Erro ao definir o tipo de usuário para: {user_name}")
                        log_error(client_name, f"Erro ao definir o tipo de usuário para: {user_name}")

                    time.sleep(2)  # Aguarda o processamento
                except:
                    print('Erro ao abrir a URL de criação de usuário.')
                    log_error(client_name, 'Erro ao abrir a URL de criação de usuário.')

    except Exception as e:
        print('Erro ao configurar o navegador ou carregar a planilha.')
        log_error(client_name, 'Erro ao configurar o navegador ou carregar a planilha.')