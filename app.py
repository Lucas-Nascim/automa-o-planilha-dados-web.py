## Importando bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

#### PROBLEMA ####

## Uso uma planilha em Excel para alimenta um sistema web. A planilha é alimentada com os dados das consultas que eu faço no sistema e isso ajuda muito na hora de organizar as informações.
## O problema seria ter que ficar inserindo os dados manualmente toda vez, mas acho que da para automatizar esse processo, deixando a tarefa mais fácil.

#### EXEMPLO ####

## Acessar o site
driver = webdriver.Chrome()
driver.get('https://www.novaliderinformatica.com.br/computadores-gamers')
    # print("Site acessado com sucesso!")

## Extrair todos os titulos
titulos =  driver.find_elements(By.XPATH,"//a[@class='nome-produto']")
    # for titulo in titulos:
    #     print(titulo.text)
    # print("Titulos encontrados com sucesso, veja acima!")

## Extrair todos os preços
precos =  driver.find_elements(By.XPATH,"//strong[@class='preco-promocional']")

## criando a planilha para guardar as informações
workbook = openpyxl.Workbook()
## Criando a aba produtos
workbook.create_sheet('produtos')
## Selecionando a aba produtos
sheet_produtos =  workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'

## Inserir todos os titulos e preços na planilha
# Usaremos o ZIP para passar por varias listas, porem ele so continua fazendo essa interação enquanto tiver itens na duas lista, ou seja no momento que ele perceber não tem mais itens nas listas ele para.
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text,preco.text])
    
workbook.save('produtos.xlsx')
print("Dados atualizados com sucesso, para visualiza-los abra a planilha produtos.xlsx")