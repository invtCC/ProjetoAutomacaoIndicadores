

"""

    Automação de Sistemas para compras de mercadorias de uma empresa


        Objetivo do Projeto.

    - Para controlar os custos, todos os dias, seu "chefe" pede um relatório com todas as compras de mercadorias da empresa. Seu trabalho é enviar um e-mail para ele assim que começar a trabalhar, com total gasto, a quantidade de produtos comprados e o preço médio dos produtos.

    
    - Essa base é ficticia, e apenas para uso de estudo e portifólio.

"""

"""
    Passo a passo.

    Passo 1 - Entrar no sistema da empresa (no link)
    
    Passo 2 - Fazer login

    Passo 3 - Exportar a base de dados

    Passo 4 - Calcular os indicadores

    Passo 5 - Enviar um e-mail para o meu chefe
"""


# Importações 

import pyautogui as py
import pyperclip
import pandas as pd
import time


# --------------------------------------------------------------------------------

    # Passo 1 - Entrando no sistema da empresa

py.PAUSE = 1

py.press('win')

py.write('Brave')

py.press('enter')   # Abre o navegador

py.hotkey('Alt', 'space')   # Abre as opções para maximizar a tela

py.press('x')   # Maximiza a tela

link = pyperclip.copy('https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema')

py.hotkey('ctrl', 'v')  # Copia o endereço do sistema da empresa

py.press('enter')   # Entra no sistema da empresa

time.sleep(2)


# --------------------------------------------------------------------------------

    # Passo 2 - Fazer login no sistema da empresa

py.click(838, 356)  # Clica no campo de login

py.write('Meu Login')   # Escreve o login

py.press('tab')   # Pula pro campo de senha

py.write('Minha senha123')  # Escreve a senha

py.click(940, 503)  # Clica em Entrar


# --------------------------------------------------------------------------------


    # Passo 3 - Exportar a base de dados

time.sleep(3)

py.click(531, 328)  # Clica na base de dados

py.click(514, 233)  # Clica em baixar 

time.sleep(3)

py.click(377, 47)   # Clica para preencher o caminho aonde você quer salvar a base

py.press('backspace')   # Apaga o que já estiver escrito

pyperclip.copy(r'C:\Users\...)   # Copia o caminho

py.hotkey('ctrl', 'v')  # Cola o caminho na barra de endereço

py.press('enter')   # Vai até o local desejado para o salvamento da base

py.click(494, 477)  # Clica em salvar


# --------------------------------------------------------------------------------

    # Passo 4 - Importar a base de dados e calcular os indicadores

base_dados = pd.read_csv('Compras.csv', sep=';')

# Total gasto

total_gasto = base_dados['ValorFinal'].sum()

# Quantidade de produtos comprados

quantidade_produtos_comprados = base_dados['Quantidade'].sum()

# Preço médio dos produtos

preco_medio = total_gasto / quantidade_produtos_comprados


# --------------------------------------------------------------------------------

    # Passo 5 - Enviar o e-mail para o chefe

py.hotkey('ctrl', 't')  #abre uma nova guia

pyperclip.copy('https://mail.google.com')   # Copia o link do email

py.hotkey('ctrl', 'v')  # Entra no e-mail

py.press('enter')

time.sleep(2)   # Espera o site carregar 
    
py.click(72, 175)   # Clica em escrever um novo e-mail

py.click(1302, 486) # Clica no campo de destinatário 

py.write('emaildoseuchefe@gmail.com')    # Escreve o destinatário

py.press('tab') # Escolhe o destinatário 

py.press('tab') # Pula para o campo de assunto 

pyperclip.copy('Relatório de Vendas')

py.hotkey('ctrl', 'v')  # Escreve o campo de assunto

py.press('tab')  # Pula para o corpo do e-mail

texto_email = f'''
Prezado Chefe,

Segue o relatório de compras.

Total Gasto: R${total_gasto:,.2f}
Quantidade de Produtos: {quantidade_produtos_comprados:,}
Preço Médio: R${preco_medio:,.2f}

Att., Fulano

'''

pyperclip.copy(texto_email)

py.hotkey('ctrl', 'v')  # Cola o e-mail no corpo do e-mail

py.hotkey('ctrl', 'enter')  # Envia o e-mail