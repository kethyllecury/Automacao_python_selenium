# Automação de Relatório de Vendas - Projeto Caneca Fina

Este repositório contém um script de automação desenvolvido para facilitar a coleta, processamento e organização de relatórios diários de vendas. A solução utiliza o Selenium para a navegação automatizada no portal web da plataforma de relatórios, juntamente com o Pandas para processamento e manipulação de dados em arquivos Excel. O objetivo principal é gerar relatórios de vendas de forma automatizada, combinando dados de diferentes fontes e organizando-os em pastas conforme o mês.

# Tecnologias Utilizadas
Python 3.x: Linguagem de programação para o desenvolvimento do script.

Selenium: Biblioteca para automação de navegação no navegador, utilizada para interagir com o portal de relatórios.

Pandas: Biblioteca para manipulação de dados em formato de tabelas (DataFrames), utilizada para editar, limpar e combinar planilhas.

webdriver_manager: Biblioteca para facilitar o gerenciamento automático do driver do navegador (Edge Chromium).

Datetime & Time: Módulos para manipulação de datas e controle de tempo durante o processo de automação.

# Funcionalidades
O script automatiza as seguintes etapas do processo de coleta e organização de relatórios:

# Login Automático
O script realiza o login automaticamente no portal de relatórios utilizando as credenciais fornecidas (e-mail e senha).

# Geração do Relatório
Após o login, o script navega até a seção de relatórios do portal e gera um relatório de vendas do dia anterior. O relatório é personalizado para exibir informações específicas, como preço, custo, número do caixa e tipo de emissor fiscal.

# Download do Relatório
O relatório gerado é exportado automaticamente para o formato Excel e baixado para o diretório de downloads do usuário.

# Processamento e Edição de Dados
Após o download, o script utiliza o Pandas para:

Carregar o arquivo Excel gerado.

Inserir a data do relatório na planilha.

Limpar dados desnecessários, removendo linhas sem dados ou com códigos inválidos.

Renomear colunas para garantir consistência entre as planilhas.

# Concatenação de Planilhas
O script carrega uma planilha pré-existente e concatena seus dados com o relatório de vendas gerado. A combinação dos dados é feita de forma a manter a consistência e evitar duplicações.

# Organização de Arquivos
O script organiza os arquivos gerados em pastas específicas, de acordo com o mês atual. Caso a pasta do mês seguinte não exista, ela é criada automaticamente. O arquivo gerado é movido para a pasta correspondente, com base na data da execução.

# Deleção de Arquivo Original
Após a movimentação do arquivo para o destino correto, o script deleta o arquivo original de entrada (Listagem.xlsx), garantindo que o diretório de downloads permaneça limpo.

# Criação Automática de Pastas
O script verifica se as pastas necessárias para armazenar os arquivos existem. Caso não existam, ele cria automaticamente as pastas, incluindo a criação de novas pastas para o mês seguinte, se necessário.

# Criptografia 
O projeto agora utiliza um ambiente virtual .venv para isolar as dependências do projeto e garantir um ambiente de desenvolvimento mais seguro e eficiente. Isso impede conflitos com outras bibliotecas e facilita a instalação das dependências de forma controlada.

# Captura de Datas de Fim de Semana
O script foi atualizado para identificar e capturar também os dados dos fins de semana. Ele realiza a coleta de relatórios de vendas tanto para os dias úteis quanto para o sábado e domingo, assegurando que o relatório cubra todo o período de vendas relevante.
