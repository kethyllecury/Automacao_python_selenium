import os
import time
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# Obtém o dia anterior
dia_anterior = datetime.now() - timedelta(days=1)
data_anterior = dia_anterior.strftime("%d-%m-%Y")
hoje = datetime.today()
mes_atual = datetime.today().strftime("%m-%Y")

pasta_downloads = r"C:\Users\sigab\Downloads"
pasta_destino = f"C:\\Users\\sigab\\OneDrive\\Driver - BPO\\Caneca Fina\\Fechamentos\\{mes_atual}\\dia a dia"
pasta_base = r"C:\Users\sigab\OneDrive\Driver - BPO\Caneca Fina\Fechamentos"
caminho_outra_planilha = fr"C:\Users\sigab\OneDrive\Driver - BPO\Caneca Fina\Fechamentos\{mes_atual}\dia a dia\{data_anterior}.xlsx"
caminho_arquivo = r"C:\Users\sigab\OneDrive\Driver - BPO\Caneca Fina\Fechamentos\12-2024\Vendas_Atualizado.xlsx"

arquivo_original = os.path.join(pasta_downloads, "Listagem.xlsx")
arquivo_novo = os.path.join(pasta_destino, f"{data_anterior}.xlsx")

def configurar_driver():
    edge_options = webdriver.EdgeOptions()
    prefs = {
        "download.default_directory": r"C:\Users\sigab\Downloads", 
        "download.prompt_for_download": False, 
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    edge_options.add_experimental_option("prefs", prefs)
    edge_options.add_argument("--start-maximized")
    
    driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=edge_options)
    return driver

def realizar_login(driver):
    driver.get("https://portal.netcontroll.com.br/#/auth/login")


    time.sleep(3)

    email_field = driver.find_element(By.ID, "email")  
    email_field.send_keys("") 

    password_field = driver.find_element(By.ID, "password")  
    password_field.send_keys("")  



    login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]")
    login_button.click()


    time.sleep(5)

def navegar_para_relatorios(driver):

    relatorios_link = driver.find_element(By.XPATH, "//a[@routerlink='/relatorio']")


    relatorios_link.click()


    driver.implicitly_wait(10)

def gerar_relatorio_vendas(driver, data_anterior):

    vendas_link = driver.find_element(By.XPATH, "//a[contains(text(), '01 - Vendas')]")


    vendas_link.click()


    driver.implicitly_wait(10)

    produtos_link = driver.find_element(By.XPATH, "//a[contains(text(), '101-Produtos')]")


    produtos_link.click()


    driver.implicitly_wait(10)

    checkbox1 = driver.find_element(By.XPATH, "//label[contains(text(), 'Exibir preço e custo (Atual)')]")


    if checkbox1.is_displayed() and checkbox1.is_enabled():
        checkbox1.click()

    checkbox2 = driver.find_element(By.XPATH, "//label[contains(text(), 'Exibir Nº Caixa')]")

    if checkbox2.is_displayed() and checkbox2.is_enabled():
        checkbox2.click()

    checkbox3 = driver.find_element(By.XPATH, "//label[contains(text(), 'Exibir Tipo Emissor Fiscal')]")

    if checkbox3.is_displayed() and checkbox3.is_enabled():
        checkbox3.click()

    input_data_inicial = driver.find_element(By.XPATH, "(//dx-date-box//input[@class='dx-texteditor-input'])[1]")
    input_data_inicial.clear()
    input_data_inicial.send_keys(data_anterior)

    input_data_final = driver.find_element(By.XPATH, "(//dx-date-box//input[@class='dx-texteditor-input'])[2]")
    input_data_final.clear()
    input_data_final.send_keys(data_anterior)

    botao_excel = driver.find_element(By.XPATH, "//div[@class='dx-button-content']//i[@class='dx-icon dx-icon-export-excel-button']")


    botao_excel.click()

    time.sleep(5)

    driver.quit()

    time.sleep(5)

def mover_arquivo(data_anterior, mes_atual):

    if hoje.day == 1:
        
        mes_anterior = hoje.replace(day=1) - timedelta(days=1)

    
        proximo_mes = mes_anterior.replace(day=28) + timedelta(days=4)  
        proximo_mes = proximo_mes.replace(day=1) 

        nova_pasta = os.path.join(pasta_base, proximo_mes.strftime("%m-%Y"), "dia a dia")

    
        if not os.path.exists(nova_pasta):
            os.makedirs(nova_pasta)
            print(f"Pasta criada: {nova_pasta}")
        else:
            print("A pasta já existe.")

    # criar o arquivo no mês certo
    if os.path.exists(arquivo_original):
        print("Arquivo encontrado!")
        df = pd.read_excel(arquivo_original)
        df.insert(0, "Data", data_anterior)
        df = df[df["Código"].astype(str).str.isnumeric()]
        df = df.dropna(subset=['Nome']) 

        print(df)
    
        df.to_excel(arquivo_novo, index=False)
    else:
        print("Arquivo NÃO encontrado. Verifique o caminho.")

    time.sleep(5)

def deletar_arquivo():
    if os.path.exists(arquivo_original):
        os.remove(arquivo_original)
        print("Arquivo Listagem.xlsx deletado com sucesso.")
    else:
        print("O arquivo Listagem.xlsx não foi encontrado.")

def editar_planilhas(caminho_arquivo, caminho_outra_planilha):

    global planilha
    global vendas
    global df_outra_planilha
    global ultima_linha_df

    print("Carregando planilhas...")

    planilha = pd.read_excel(caminho_arquivo, sheet_name=None)
    print(f"Planilhas carregadas: {list(planilha.keys())}")

    vendas = planilha['Sheet1']
    print(f"Colunas da aba 'Vendas': {vendas.columns}")

    df_outra_planilha = pd.read_excel(caminho_outra_planilha, sheet_name='Sheet1')
    print(f"Colunas da segunda planilha antes do renomeio: {df_outra_planilha.columns}")

    print("Exemplo de datas convertidas em 'Vendas':")
    print(vendas[['Data Venda']].head())

    print("Exemplo de datas convertidas na segunda planilha:")
    print(df_outra_planilha[['Data']].head())

    df_outra_planilha = df_outra_planilha.rename(columns={'Data': 'Data Venda', 'Valor Total': 'Valor', 'Emissor Fiscal': 'Tipo Cfe'})

    print(f"Colunas da segunda planilha após renomeio: {df_outra_planilha.columns}")

    ultima_linha_df = vendas['Data Venda'].last_valid_index()
    print(f"Última linha válida na aba 'Vendas': {ultima_linha_df}")

def concatenar_planilhas(vendas, planilha, df_outra_planilha, ultima_linha_df):

    df_concatenado = pd.concat([vendas.iloc[:ultima_linha_df+1], df_outra_planilha], ignore_index=True)
    print(f"Tamanho do dataframe antes: {len(vendas)}, tamanho do dataframe depois da concatenação: {len(df_concatenado)}")

    planilha['Sheet1'] = df_concatenado

    df_concatenado.to_excel(caminho_arquivo, index=False)
    print(f"Arquivo salvo em: {caminho_arquivo}")



 
driver = configurar_driver()

realizar_login(driver)

navegar_para_relatorios(driver)

gerar_relatorio_vendas(driver, data_anterior)

mover_arquivo(data_anterior, mes_atual)

deletar_arquivo()

editar_planilhas(caminho_arquivo, caminho_outra_planilha)

concatenar_planilhas(vendas, planilha, df_outra_planilha, ultima_linha_df)











