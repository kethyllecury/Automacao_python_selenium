# Automação de Relatório de Vendas - Projeto Caneca Fina

Este repositório contém um script de automação desenvolvido para facilitar a coleta, processamento e organização de relatórios diários de vendas. O script utiliza o Selenium para navegar no portal web da plataforma de relatórios e o Pandas para processar e manipular dados em arquivos Excel. O objetivo principal é gerar relatórios de vendas de forma automatizada, combinando dados de diferentes fontes e organizando-os em pastas conforme o mês.

# Tecnologias Utilizadas

Python 3.x: Linguagem de programação utilizada para o desenvolvimento do script.

Selenium: Biblioteca para automação de navegação no navegador, utilizada para interagir com o portal de relatórios.

Pandas: Biblioteca poderosa para manipulação de dados em formato de tabelas (DataFrames), utilizada para editar, limpar e combinar planilhas.

webdriver_manager: Biblioteca para facilitar o gerenciamento automático do driver do navegador (Edge Chromium).

Datetime & Time: Módulos para manipulação de datas e controle de tempo no processo de automação.

# Funcionalidades

O script tem várias funcionalidades automatizadas para facilitar o processo de obtenção e organização dos relatórios:

1. Login Automático
O script faz login automaticamente no portal de relatórios com as credenciais fornecidas (e-mail e senha).

2. Geração de Relatório
Após o login, o script navega até a seção de relatórios do portal e gera um relatório de vendas para o dia anterior. Ele personaliza o relatório, selecionando as opções específicas como exibir preço, custo, número do caixa e tipo de emissor fiscal.

3. Download do Relatório
O relatório gerado é automaticamente exportado para o formato Excel e baixado para o diretório de downloads do usuário.

4. Processamento e Edição de Dados
Após o download do relatório, o script utiliza a biblioteca Pandas para:

 4.1 Carregar o relatório em formato Excel.
 
 4.2 Inserir a data do relatório na planilha.
 
 4.3 Limpar dados desnecessários e garantir que a planilha contenha apenas dados válidos (por exemplo, remove linhas sem dados ou com códigos inválidos).
 
 4.4 Renomear colunas para garantir consistência entre as planilhas.

5. Concatenação de Planilhas:
   
O script carrega outra planilha pré-existente e concatena seus dados com o relatório de vendas gerado. Os dados são combinados de forma a manter a consistência e evitar duplicação.

6. Organização de Arquivos:
   
Os arquivos são organizados em pastas de acordo com o mês atual. Caso a pasta do mês seguinte não exista, o script cria a pasta automaticamente. O arquivo gerado é movido para a pasta correta, com base no mês da execução.

7. Deleção de Arquivo Original:
    
Após a movimentação do arquivo para o destino correto, o script deleta o arquivo original de entrada (Listagem.xlsx), garantindo que o diretório de downloads seja limpo.

8. Criação Automática de Pastas:
   
O script garante que as pastas necessárias para armazenar os arquivos estejam sempre criadas, incluindo a criação de novas pastas para o mês seguinte, caso necessário.
