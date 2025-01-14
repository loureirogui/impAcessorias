Automacao de Cadastro e Gerenciamento Tributário

Este projeto realiza tarefas de automação para o sistemas Acessórias, incluindo o cadastro de usuários, criação e alocação de obrigações, regimes tributários, cadastro de empresas e contatos.

Tecnologias Utilizadas

Linguagem: Python

Interface Gráfica: tkinter

Automatização de Navegação: selenium

Manipulação de Planilhas: openpyxl

Funcionalidades

Cadastro de Usuários:

Automatiza o cadastro de novos usuários no sistema.

Gerenciamento de Obrigações:

Atualização de obrigações.

Cria e aloca obrigações nos regimes tributários.

Cadastro de Empresas:

Registra novas empresas e seus respectivos contatos e regimes tributários.

Configuração e Execução

Requisitos

Python 3.8+

Navegador Microsoft Edge

Instalando Dependências

Certifique-se de ter o pip instalado e execute o seguinte comando para instalar as dependências:

pip install -r requirements.txt


Configurando o WebDriver

O projeto utiliza o webdriver-manager para gerenciar automaticamente o driver do Edge. Certifique-se de que o navegador esteja atualizado.


Configuração do Arquivo Excel

O projeto utiliza uma planilha Excel para entrada de dados. Um modelo está incluído no diretório do projeto: Planilha modelo Acessórias.xlsx.
Abaixo um vídeo explicativo sobre como preencher a planilha:
https://drive.google.com/file/d/1O_lrB4qNIhpPeFVy8JA8dFk6VvRuRQxf/view?usp=drive_link



project/
├── components/             # Módulos do sistema
│   ├── createUser.py     # Cadastro de usuários
│   ├── obrigacao.py      # Gerenciamento de obrigações
│   ├── updateTax.py     # Atualização de regimes tributários
│   └── createCompany.py  # Cadastro de empresas
├── Planilha modelo Acessórias.xlsx # Planilha modelo para importação
├── main.py                # Arquivo principal
├── requirements.txt      # Dependências do projeto
└── README.md             # Documentação do projeto


