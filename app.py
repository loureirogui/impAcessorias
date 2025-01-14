import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from components.createUser import create_users
from components.obrigacao import atualizaObrigacao
from components.uptadeTax import update_tax_regime
from components.createCompany import register_company


def anexar_planilha(): # Função para selecionar a planilha
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_planilha.delete(0, tk.END)  # Limpa qualquer entrada anterior
        entry_planilha.insert(0, file_path)


def enviar_dados(): # Função para coletar os dados do formulário
    nomeCliente = entry_cliente.get()
    emailLogin = entry_email.get()
    senhaLogin = entry_senha.get()
    nomePlanilha = entry_planilha.get()
    
    if not nomeCliente or not emailLogin or not senhaLogin or not nomePlanilha: # Verifica se não tem campos necessários vazios
        messagebox.showwarning("Campos vazios", "Por favor, preencha todos os campos.")
        return
    
    
    if var_criarUsuarios.get(): # Verifica quais checkboxes estão selecionados e executa as funções correspondentes
        create_users(nomeCliente, emailLogin, senhaLogin, nomePlanilha)
    
    if var_atualizaObrigacao.get():
        atualizaObrigacao(nomeCliente, emailLogin, senhaLogin, nomePlanilha)
    
    if var_atualizaRegime.get():
        update_tax_regime(nomeCliente, emailLogin, senhaLogin, nomePlanilha)
    
    if var_cadastraEmpresa.get():
        register_company(nomeCliente, emailLogin, senhaLogin, nomePlanilha)


root = tk.Tk() # Criação da janela principal
root.title("Automação de implantação")
root.geometry("400x600") # Tamanho da janela



lbl_cliente = tk.Label(root, text="Nome do Escritório:") # Campo para digitar Nome da contabilidade
lbl_cliente.pack(pady=5)
entry_cliente = tk.Entry(root, width=50)
entry_cliente.pack(pady=5)


lbl_email = tk.Label(root, text="Email de Login:") # Campo para digitar email de login Acessorias
lbl_email.pack(pady=5)
entry_email = tk.Entry(root, width=50)
entry_email.pack(pady=5)


lbl_senha = tk.Label(root, text="Senha de Login:") # Campo para digitar senha de login Acessorias
lbl_senha.pack(pady=5)
entry_senha = tk.Entry(root, show="*", width=50)
entry_senha.pack(pady=5)


lbl_planilha = tk.Label(root, text="Planilha do Cliente:") # Campo para anexar a planilha
lbl_planilha.pack(pady=5)


entry_planilha = tk.Entry(root, width=50) # Campo que mostrará o caminho da planilha
entry_planilha.pack(pady=5)


btn_anexar = tk.Button(root, text="Anexar Planilha", command=anexar_planilha) # Botão para anexar a planilha
btn_anexar.pack(pady=5)


var_criarUsuarios = tk.BooleanVar()  # Variáveis de controle para os checkboxes de quais funções chamar
var_atualizaObrigacao = tk.BooleanVar()
var_atualizaRegime = tk.BooleanVar()
var_cadastraEmpresa = tk.BooleanVar()

# Checkboxes para selecionar as funções a serem executadas
chk_criarUsuarios = tk.Checkbutton(root, text="Criar Usuários", variable=var_criarUsuarios)
chk_criarUsuarios.pack(pady=5)

chk_atualizaObrigacao = tk.Checkbutton(root, text="Atualizar Obrigações", variable=var_atualizaObrigacao)
chk_atualizaObrigacao.pack(pady=5)

chk_atualizaRegime = tk.Checkbutton(root, text="Atualizar Regime", variable=var_atualizaRegime)
chk_atualizaRegime.pack(pady=5)

chk_cadastraEmpresa = tk.Checkbutton(root, text="Cadastrar Empresa", variable=var_cadastraEmpresa)
chk_cadastraEmpresa.pack(pady=5)

# Botão para enviar os dados
btn_enviar = tk.Button(root, text="Enviar", command=enviar_dados)
btn_enviar.pack(pady=20)

# Inicia o loop principal da interface
root.mainloop()