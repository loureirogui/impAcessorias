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
import openpyxl
import unicodedata
import random
import re
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def register_company(client_name, login_email, login_password, spreadsheet_name):
    driver_path = 'msedgedriver.exe'# Caminho para o driver do Edge

    
    edge_options = Options()# Configurações do Edge
    edge_options.add_argument('--log-level=3')  # Define o nível de log como 'fatal', suprimindo a maioria das mensagens de erro.
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Exclui certos logs.
    edge_driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=edge_options)# Inicializa o navegador Edge
    edge_driver.set_window_size(1300, 800) # Tamanho da tela pré-definido para evitar erros de XPATH
    url = f"https://app.acessorias.com" # Página de login
    edge_driver.get(url)

    
    try:# Lógica de login
        # Espera o campo de e-mail aparecer
        email_input = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'mailAC'))
        )
        # Insere o e-mail no campo
        email_input.send_keys(login_email)
    except Exception:
        print("Erro ao inserir o e-mail no campo de login:")
    
    try:
        # Espera o campo de senha aparecer
        password_input = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'passAC'))
        )
        # Insere a senha no campo
        password_input.send_keys(login_password)
    except Exception:
        print("Erro ao inserir a senha no campo de senha:")
    

    # Espera o botão de login aparecer
    try:
        login_button = WebDriverWait(edge_driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.button.rounded.large.expanded.primary-degrade.btn-enviar'))
        )
        # Clica no botão de login
        login_button.click()
    except Exception:
        print("Erro ao clicar no botão de login:")
    
    time.sleep(2)

    # Carrega o arquivo .xlsx
    workbook = openpyxl.load_workbook(spreadsheet_name)

    # Seleciona a planilha
    sheet = workbook['Empresas']

    possible_numbers = list(range(60000, 64001))

    # Função para obter um ID aleatório caso o do cliente esteja em uso
    def get_random_id():
        if possible_numbers:  # Verifica se ainda há números disponíveis
            return possible_numbers.pop(random.randint(0, len(possible_numbers) - 1))
        else:
            raise ValueError("Todos os IDs possíveis já foram usados.")

    for row in sheet.iter_rows(min_row=3, min_col=1, max_col=11):
        
        company_id = row[0].value  # Coluna 1: Código do Sistema Contábil (ID)
        trade_name = row[1].value  # Coluna 2: Nome Fantasia
        legal_name = row[2].value  # Coluna 3: Razão Social
        cnpj = row[3].value        # Coluna 4: CNPJ/CPF/CEI/CAEPF
        state_registration = row[4].value  # Coluna 5: Insc. Estadual
        state_registration_uf = row[5].value  # Coluna 6: Insc. Estadual UF
        contact_name = row[6].value  # Coluna 7: Nome do Contato
        contact_email = row[7].value  # Coluna 8: E-mail do Contato
        tax_regime = row[8].value  # Coluna 9: Regime Tributário
        nickname = row[9].value  # Coluna 11: Apelido e-Contínuo
        contact_phone = row[10].value
        if cnpj:
            try:
                url = 'https://app.acessorias.com/sysmain.php?m=105&act=a' # URL para adicionar uma nova empresa
                edge_driver.get(url)
            except Exception as e:
                print("Erro ao acessar página de adicionar empresa")
               
            
            try:# Inserir CNPJ no campo
                # Espera o campo de CNPJ aparecer
                cnpj_input = WebDriverWait(edge_driver, 10).until(
                    EC.visibility_of_element_located((By.NAME, 'field_EmpCNPJ'))
                )
                # Insere o CNPJ no campo
                cnpj_input.send_keys(cnpj)
                cnpj_input.send_keys(Keys.TAB) #Ao mandar o TAB dispara o gatilho caso o CPF/CNPJ seja inválido. Evitando o erro de carregamento e leitura do erro
            except:
                print("Erro ao inserir o CNPJ da empresa")
            time.sleep(0.5)
            try:
                # Se o botão "btCNPJ" não for encontrado, tenta localizar o botão com o ID "btCPF"
                searchButton = WebDriverWait(edge_driver, 1).until(
                    EC.element_to_be_clickable((By.ID, 'btCNPJ'))
                )
                searchButton.click()
            except:
                try:
                    # Se o botão "btCNPJ" não for encontrado, provavelmente o Acessórias não reconheceu o número como um identificador válido"
                    searchButton = WebDriverWait(edge_driver, 1).until(
                        EC.element_to_be_clickable((By.ID, 'btCPF'))
                    )
                    searchButton.click()
                except:
                    print(f"CNPJ/CPF: {cnpj} inválido, cadastrando mesmo assim.")
                
            time.sleep(1) # Sem o essa pausa de 1s dá erro de carregamento no pop up de erro
            try:
                # Aguarda até que o pop-up esteja visível (tempo máximo de 5 segundos)
                popup_error = WebDriverWait(edge_driver, 1).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, '.swal2-popup.swal2-modal.swal2-show'))
                )
                # Procura pelo botão "Acessar o cadastro" dentro do pop-up e clica nele
                access_registration_button = popup_error.find_element(By.CSS_SELECTOR, '.swal2-confirm.btn.btn-success.marginZ')
                access_registration_button.click()
                time.sleep(1)
            except Exception as e:
                pass
            
            try: # Pop up de CPf/CNPJ inválido, infelizmente o desenvolvedor duplicou o elemento html, então tive que lidar com isso duplicando a função kkkkkk
                # Aguarda até que o pop-up esteja visível
                popup_error = WebDriverWait(edge_driver, 1).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, '.swal2-popup.swal2-modal.swal2-show'))
                )
                # Procura pelo botão "Continuar" no pop-up e clica nele
                continue_button = popup_error.find_element(By.CSS_SELECTOR, '.swal2-confirm.btn.btn-danger.marginZ')
                continue_button.click()
            except Exception as e:
                pass

            time.sleep(0.5)

            try: # Pop up de CPf/CNPJ inválido, infelizmente o desenvolvedor duplicou o elemento html, então tive que lidar com isso duplicando a função kkkkkk
                # Aguarda até que o pop-up esteja visível
                popup_error = WebDriverWait(edge_driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, '.swal2-popup.swal2-modal.swal2-show'))
                )
                # Procura pelo botão "Continuar" no pop-up e clica nele
                continue_button = popup_error.find_element(By.CSS_SELECTOR, '.swal2-confirm.btn.btn-danger.marginZ')
                continue_button.click()
            except Exception as e:
                pass
            
            time.sleep(0.5)

            # Função para normalizar o texto (remover acentos, caracteres especiais, e deixar em maiúsculas)
            def normalize_text(text):
                text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
                text = text.replace('/', '')  # Substitui '/' por vazio O Acessórias remove "/" do título dos regimes automaticamente, então também preciso fazer isso aqui para comparar
                text = text.upper()
                return text

           
            try: # Regime tributário
                # Encontra o seletor pelo nome do regime
                select_tax_regime = WebDriverWait(edge_driver, 2).until(
                    EC.visibility_of_element_located((By.NAME, 'field_EmpRegID'))
                )

                select = Select(select_tax_regime)
                # Normaliza o texto do regime para comparação
                normalized_tax_regime = normalize_text(tax_regime)
                found = False
                for option in select.options:
                    # Normaliza o texto da opção
                    option_text_normalized = normalize_text(option.text)
                    if option_text_normalized == normalized_tax_regime:
                        select.select_by_visible_text(option.text)
                        found = True
                        time.sleep(0.5) # Tempo para o modal de regime selecionado e obrigações alocadas apareça 
                        try: # Clicar no "OK"
                            ok_button = WebDriverWait(edge_driver, 3).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                            )
                            ok_button.click()
                        except Exception:
                            pass
                        break
                
                
                # Aguarda o botão "OK" aparecer e clicar nele
                try:
                    ok_button = WebDriverWait(edge_driver, 3).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                    )
                    ok_button.click()
                except Exception:
                    pass
                
                if not found:
                    print(f"A empresa de CNPJ: {cnpj} foi cadastrada sem Regime tributário pois o '{tax_regime}' não está de acordo com o apontado na planilha")

            except Exception as e:
                pass
            
            time.sleep(1)
            
            if company_id: # Inserir o ID da empresa
                try:
                    # Espera o campo de ID da empresa aparecer
                    id_input = WebDriverWait(edge_driver, 3).until(
                        EC.visibility_of_element_located((By.NAME, 'EmpNewID'))
                    )
                    id_input.clear()
                    id_input.send_keys(company_id)
                    id_input.send_keys(Keys.TAB) # TAB ativa o gatilho para verificar se o id já está em uso

                    
                    try: # Verifica se o ID está em uso, se sim, cria id aleatório
                        alert_element = WebDriverWait(edge_driver, 2).until(
                            EC.visibility_of_element_located((By.CSS_SELECTOR, '.swal2-popup.swal2-modal.swal2-show'))
                        )
                        if alert_element: 
                            ok_button = WebDriverWait(edge_driver, 3).until(
                                EC.visibility_of_element_located((By.XPATH, '//*[@id="body"]/div[7]/div/div[3]/button[1]'))
                            )
                            ok_button.click()
                            random_id = get_random_id()
                            id_input.clear()
                            id_input.send_keys(random_id)
                            print(f"A empresa de CNPJ: {cnpj} terá um id aleatório de número {random_id} pois o id {company_id} da planilha já está em uso.")
                        else:
                            pass
                    except Exception:
                        pass
                except Exception as e:
                    pass

            time.sleep(0.5)
            
            if trade_name:# Inserir nome fantasia ou razão social (caso o cliente não preencha um dos campos, posso passar outro argumento para não deixar vazio evitando erros do sistema)
                try:
                    trade_name_input = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, 'field_EmpFantasia'))
                    )
                    trade_name_input.clear()
                    trade_name_input.send_keys(trade_name)
                except Exception as e:
                    pass
            else:
                try:
                    trade_name_input = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, 'field_EmpFantasia'))
                    )
                    trade_name_input.clear()
                    trade_name_input.send_keys(legal_name)
                except Exception as e:
                    pass
            
            # Inserir razão social ou nome fantasia (caso o cliente não preencha um dos campos, posso passar outro argumento para não deixar vazio evitando erros do sistema)
            if legal_name:
                try:
                    legal_name_input = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, 'field_EmpNome'))
                    )
                    legal_name_input.clear()
                    legal_name_input.send_keys(legal_name)
                except Exception as e:
                    print("Erro ao inserir a razão social.")
            else:
                try:
                    legal_name_input = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, 'field_EmpNome'))
                    )
                    legal_name_input.clear()
                    legal_name_input.send_keys(trade_name)
                except Exception as e:
                    print("Erro ao inserir a razão social.")

            if nickname:    
                
                try:# Localiza e insere o Apelido eContínuo caso possua
                    nickname_input = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="EmpApelido"]'))
                    )
                    nickname_input.clear()
                    nickname_input.send_keys(nickname)
                except:
                    print("Error inserting e-Contínuo nickname")

            time.sleep(0.5)
            
            if state_registration: # Localiza e insere número de inscrição, UF e Clica do botão de adicionar
                try:
                    adress_icon_element = WebDriverWait(edge_driver, 10).until(
                        EC.element_to_be_clickable((By.ID, 'iDivEnd'))
                    )
                    # Se estiver com o botão de Inscrições e Endereços Cinza, clica nele para poder editar os campos
                    adress_icon_class = adress_icon_element.get_attribute("class")
                    if 'grey' in adress_icon_class:
                        edge_driver.execute_script("arguments[0].scrollIntoView(true);", adress_icon_element)
                        adress_icon_element.click()
                except:
                    pass

                try:
                    state_registration_input = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, 'EmpNewIE'))
                    )
                    state_registration_input.send_keys(state_registration)
                    
                    try:
                        select_uf = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'EmpIEUF'))
                        )
                        select = Select(select_uf)
                        
                        # Processo de normalização para comparar os UFS e selecionar a correspondente
                        normalized_uf = state_registration_uf.upper()
                        for option in select.options:
                            if option.text == normalized_uf:
                                select.select_by_visible_text(option.text)
                                time.sleep(0.3)
                                edge_driver.execute_script("addIE();")
                                break
                            else:
                                pass                     
                    except Exception as e:
                        pass
                except Exception as e:
                    print("Error processing company state registration.")

            time.sleep(0.5)
                
            if contact_name and contact_email:# Divide os emails por '/' ou ";"
                
                contact_names = re.split(r'[/;,]', contact_name)
                contact_emails = re.split(r'[/;,]', contact_email)

                # Itera sobre os contatos adicionando eles
                for i, (name, email) in enumerate(zip(contact_names, contact_emails)):
                    name = name.strip()
                    email = email.strip()
                    
                    try:
                        contact_icon_element = WebDriverWait(edge_driver, 10).until(
                            EC.element_to_be_clickable((By.ID, 'iDivCtt'))
                        )

                        # Se o botão de contatos estiver desabilitado, habilita ele para mostrar o campo
                        contact_icon_class = contact_icon_element.get_attribute("class")
                        if 'grey' in contact_icon_class:
                            edge_driver.execute_script("arguments[0].scrollIntoView(true);", contact_icon_element)
                            contact_icon_element.click()
            
                        contact_name_input = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'CttNome_0'))
                        )
                        contact_name_input.clear()
                        contact_name_input.send_keys(name)
                        
                        contact_email_input = WebDriverWait(edge_driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, 'CttEMail_0'))
                        )
                        contact_email_input.clear()
                        contact_email_input.send_keys(email)
                        
                        if contact_phone: # Insere o telefone
                            contact_phone_input = WebDriverWait(edge_driver, 10).until(
                                EC.visibility_of_element_located((By.NAME, 'CttFone_0'))
                            )
                            contact_phone_input.clear()
                            contact_phone_input.send_keys(contact_phone)

                        time.sleep(0.5)
                    
                        try:
                            # Marcar todos os departamentos para os contatos receberem correspondência das entregas
                            mark_department_button = WebDriverWait(edge_driver, 2).until(
                                EC.visibility_of_element_located((By.XPATH, '//*[@id="dptoCtt_New_0"]/div[1]/div[1]/span/a[1]'))
                            )
                            mark_department_button.click()
                        except Exception as e:
                            pass

                        # Salvar o contato
                        try:
                            edge_driver.execute_script("addCtt('0', true);")
                            time.sleep(0.5)
                        except Exception as e:
                            print(f"Erro ao salvar o contato {name} da empresa {cnpj}.")
                        
                        
                        try: # Lida com potenciais mensagens de erro no contato
                            modal_element = WebDriverWait(edge_driver, 2).until(
                                EC.visibility_of_element_located((By.CLASS_NAME, 'swal2-popup'))
                            )
                            if modal_element:
                                ok_button = WebDriverWait(edge_driver, 2).until(
                                    EC.element_to_be_clickable((By.CLASS_NAME, 'swal2-confirm'))
                                )
                                ok_button.click()
                                print(f"Erro ao salvar o contato {name} da empresa {cnpj}.")
                        except Exception as e:
                            pass
    
                    except Exception as e:
                        print(f"Erro ao inserir o contato {name}: {e}")
                        
            try: #Executa botão de salvar
                edge_driver.execute_script("check_form(this);")
            except:
                print(f"Empresa de CPF/CNPJ: {cnpj} não cadastrada")
            time.sleep(2)
        else:
            break

    edge_driver.quit()

