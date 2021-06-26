# Base de Dados: https://drive.google.com/drive/folders/1o2lpxoi9heyQV1hIlsHXWSfDkBPtze-V?usp=sharing
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

# Para rodar o chrome em 2º plano
# from selenium.webdriver.chrome.options import Options
# chrome_options = Options()
# chrome_options.headless = True # also works
# nav = webdriver.Chrome(options=chrome_options)

nav = webdriver.Chrome()

# Pesquisar cotação dolar
nav.get('https://www.google.com/')
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dólar')
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_dolar = nav.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

# Pesquisar cotação euro
nav.get('https://www.google.com/')
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro')
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = nav.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)

# Pegar cotação ouro
nav.get('https://www.melhorcambio.com/')
aba_original = nav.window_handles[0]
nav.find_element_by_xpath('//*[@id="commodity-hoje"]/tbody/tr[2]/td[2]/a/img').click()
aba_nova = nav.window_handles[1]
nav.switch_to.window(aba_nova)
cotacao_ouro = nav.find_element_by_id('comercial').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(',', '.')
print(cotacao_ouro)

nav.quit()

# Importação da base de dados
produtos_df = pd.read_excel('Produtos.xlsx')
display(produtos_df)

# Atualização dos preços e do cálculo do Preço Final
produtos_df.loc[produtos_df['Moeda']=='Dólar', 'Cotação'] = float(cotacao_dolar)
produtos_df.loc[produtos_df['Moeda']=='Euro', 'Cotação'] = float(cotacao_euro)
produtos_df.loc[produtos_df['Moeda']=='Ouro', 'Cotação'] = float(cotacao_ouro)

produtos_df['Preço Base Reais'] = produtos_df['Cotação'] * produtos_df['Preço Base Original']
produtos_df['Preço Final'] = produtos_df['Margem'] * produtos_df['Preço Base Reais']
# produtos_df['Preço Final'] = produtos_df['Preço Final'].map('{:.2f}'.format)
display(produtos_df)

# Exportar base atualizada
produtos_df.to_excel('ProdutosAtualizados.xlsx', index=False)