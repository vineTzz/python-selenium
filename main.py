# Importar Modulos #
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from time import sleep

# Autorizar Download Automatico e Criar navegador #
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\alxia\OneDrive\Área de Trabalho\github\python + selenium\nota fiscal excel\notas baixadas",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})
navegador = webdriver.Chrome(options=options)

# Fazer Login no site #
navegador.get(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + selenium\nota fiscal excel\login.html')
navegador.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('vinetzz', Keys.TAB)
navegador.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('123456', Keys.ENTER)

# Importar Base de Dados #
arquivo_cliente = pd.read_excel(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + selenium\nota fiscal excel\NotasEmitir.xlsx')
for linha in arquivo_cliente.index:

# Preencher Formulario #
  # Nome #
  navegador.find_element(By.XPATH, '//*[@id="nome"]').send_keys(arquivo_cliente.loc[linha, 'Cliente'])

  # Endereço #
  navegador.find_element(By.NAME, 'endereco').send_keys(arquivo_cliente.loc[linha, 'Endereço'])

  # Bairro #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[3]').send_keys(arquivo_cliente.loc[linha, 'Bairro'])

  # Municipio #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[4]').send_keys(arquivo_cliente.loc[linha, 'Municipio'])

  # CEP #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[5]').send_keys(str(arquivo_cliente.loc[linha, 'CEP']))

  # UF #
  navegador.find_element(By.XPATH, '/html/body/div/form/select').send_keys(arquivo_cliente.loc[linha, 'UF'])

  # CPF #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[6]').send_keys(str(arquivo_cliente.loc[linha, 'CPF/CNPJ']))

  # Inscrição Estadual #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[7]').send_keys(str(arquivo_cliente.loc[linha, 'Inscricao Estadual']))

  # Descriçao Produto #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[8]').send_keys(arquivo_cliente.loc[linha, 'Descrição'])

  # Quantidade #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[9]').send_keys(str(arquivo_cliente.loc[linha, 'Quantidade']))

  # Valor Unitario #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[10]').send_keys(str(arquivo_cliente.loc[linha, 'Valor Unitario']))

  # Valor Total #
  navegador.find_element(By.XPATH, '/html/body/div/form/input[11]').send_keys(str(arquivo_cliente.loc[linha, 'Valor Total']))

  # Enviar Formulario #
  navegador.find_element(By.XPATH, '/html/body/div/form/button').send_keys(Keys.ENTER)

  # Recarregar Pagina #
  navegador.refresh()

  sleep(10)


