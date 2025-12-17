from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import os
import sys

# Entrando na plataforma Logshare
usuario_logshare = input("Digite seu usuário do Logshare: ")
senha_logshare = input("Digite sua senha do Logshare: ")
sys.stderr = open(os.devnull, 'w')  # Suprime mensagens de erro do Selenium

navegador = webdriver.Chrome()
time.sleep(2)
navegador.maximize_window()
navegador.get("https://logshare.app/login")
time.sleep(2)
usuario = navegador.find_element(By.XPATH, '//input[@type="text" or @name="username"]')
usuario.send_keys({usuario_logshare})
senha = navegador.find_element(By.XPATH, '//input[@type="password"]')
senha.send_keys({senha_logshare})
entrar = navegador.find_element(By.XPATH, '//button[contains(.,"Entrar")]')
entrar.click()
time.sleep(2)

base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
caminho_arquivo = os.path.join(base_dir, "Planilha de Fretes - Venda.xlsx")

if not os.path.exists(caminho_arquivo):
    raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

df = pd.read_excel(caminho_arquivo, sheet_name="Preenchimento")

for index, row in df.iterrows():
    # Cria variáveis para cada coluna
    val_tipo_origem = row['Tipo Origem']
    val_nome_origem = row['Nome Origem']
    val_tipo_destino = row['Tipo Destino']
    val_nome_destino = row['Nome Destino']
    val_tipo_veiculo = row['Tipo Veículo']
    val_categoria_veiculo = row['Categoria Veículo']
    val_frequencia = row['Frequência Mensal']
    val_frete = row['Valor Frete']

    # Acessando a página de ofertas de venda
    navegador.get("https://logshare.app/ofertas-venda")
    time.sleep(2)
    botao_nova_oferta = navegador.find_element(By.XPATH, "//div[contains(text(), 'Nova Oferta')]")
    botao_nova_oferta.click()
    time.sleep(2)

    # Preenchendo a origem do frete
    botao_tipo_origem = navegador.find_element(
    By.XPATH,
    "//label[@for='origin.type']"
    "/following-sibling::div"
    "//div[contains(@class,'indicatorContainer')]")
    botao_tipo_origem.click()
    tipo_origem = navegador.find_element(By.XPATH, f"//div[text()='{val_tipo_origem}']")
    tipo_origem.click()
    time.sleep(1)

    dropdown_origem = navegador.find_element(By.XPATH,
        "//label[@for='origin.name']"
        "/following-sibling::div"
        "//div[contains(@class,'css-1xc3v61-indicatorContainer')][last()]")
    dropdown_origem.click()
    time.sleep(1)
    input_origem = navegador.find_element(
    By.XPATH,
    "//label[@for='origin.name']"
    "/following-sibling::div"
    "//div[contains(@class,'css-19bb58m')]"
    "//input[@type='text']")
    input_origem.send_keys(f"{val_nome_origem}")
    time.sleep(1)
    input_origem.send_keys(Keys.ENTER)
    time.sleep(1)

    # Preenchendo o destino do frete
    botao_tipo_destino = navegador.find_element(
    By.XPATH,
    "//label[@for='destination.type']"
    "/following-sibling::div"
    "//div[contains(@class,'indicatorContainer')]")
    botao_tipo_destino.click()
    time.sleep(1)
    input_tipo_destino = navegador.find_element(
    By.XPATH,
    "//label[@for='destination.type']"
    "/following-sibling::div"
    "//div[contains(@class,'css-19bb58m')]"
    "//input[@type='text']")
    input_tipo_destino.send_keys(val_tipo_destino)
    time.sleep(1)
    input_tipo_destino.send_keys(Keys.ENTER)
    time.sleep(1)

    dropdown_destino = navegador.find_element(By.XPATH,
        "//label[@for='destination.name']"
        "/following-sibling::div"
        "//div[contains(@class,'css-1xc3v61-indicatorContainer')][last()]")
    navegador.execute_script(
    "arguments[0].scrollIntoView({block: 'center'});",
    dropdown_destino)
    time.sleep(1)
    dropdown_destino.click()
    time.sleep(1)

    input_destino = navegador.find_element(
    By.XPATH,
    "//label[@for='destination.name']"
    "/following-sibling::div"
    "//div[contains(@class,'css-19bb58m')]"
    "//input[@type='text']")
    input_destino.send_keys(f"{val_nome_destino}")
    time.sleep(1)
    input_destino.send_keys(Keys.ENTER)
    time.sleep(1)

    # Cadastrando a tarifa
    botao_cadastrar_tarifa = navegador.find_element(By.XPATH, "//div[contains(text(), 'Cadastrar Tarifa')]")
    botao_cadastrar_tarifa.click()
    time.sleep(1)

    # Preenchendo o tipo de veículo
    dropdown_tipo_veiculo = navegador.find_element(
    By.XPATH,
    "//label[@for='vehicleType']"
    "/following-sibling::div"
    "//div[contains(@class,'indicatorContainer')]")
    dropdown_tipo_veiculo.click()
    time.sleep(1)
    input_tipo_veiculo = navegador.find_element(
    By.XPATH,
    "//label[@for='vehicleType']"
    "/following-sibling::div"
    "//div[contains(@class,'css-19bb58m')]"
    "//input[@type='text']")
    input_tipo_veiculo.send_keys(f"{val_tipo_veiculo}")
    time.sleep(1)
    input_tipo_veiculo.send_keys(Keys.ENTER)
    time.sleep(1)

    # Preenchendo a categoria de veículo
    dropdown_categoria_veiculo = navegador.find_element(
    By.XPATH,
    "//label[@for='vehicleCategory']"
    "/following-sibling::div"
    "//div[contains(@class,'indicatorContainer')]")
    dropdown_categoria_veiculo.click()
    time.sleep(1)
    input_categoria_veiculo = navegador.find_element(
    By.XPATH,
    "//label[@for='vehicleCategory']"
    "/following-sibling::div"
    "//div[contains(@class,'css-19bb58m')]"
    "//input[@type='text']")
    input_categoria_veiculo.send_keys(f"{val_categoria_veiculo}")
    time.sleep(1)
    input_categoria_veiculo.send_keys(Keys.ENTER)
    time.sleep(1)

    #Preenchendo frequência e valor do frete
    frequencia = navegador.find_element(By.ID, "frequency")
    frequencia.click()
    time.sleep(1)
    frequencia.click()
    frequencia.send_keys(Keys.CONTROL + "a")
    frequencia.send_keys(Keys.DELETE)

    frequencia.send_keys(val_frequencia)
    frequencia.send_keys(Keys.ENTER)
    time.sleep(1)

    valor_frete = navegador.find_element(By.ID, "valueShipper")
    valor_frete.click()
    time.sleep(1)
    valor_frete.send_keys(Keys.CONTROL + "a")
    time.sleep(1)
    valor_frete.send_keys(val_frete)
    time.sleep(1)

    #Salvando a tarifa e a oferta de venda
    botao_salvar_tarifa = navegador.find_element(By.XPATH, "//div[contains(text(), 'Salvar')]")
    botao_salvar_tarifa.click()
    time.sleep(1)

    botao_salvar_oferta = navegador.find_element(By.XPATH, "//div[contains(text(), 'Salvar')]")
    botao_salvar_oferta.click()
    time.sleep(1)

    print(f"Oferta {index + 1} de {len(df)} criada com sucesso!")

print("Todas as ofertas foram criadas com sucesso!")
navegador.quit()

time.sleep(10)
