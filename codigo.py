import pyautogui as pg
import pandas as pd
import numpy as np
import openpyxl as op

link = "https://dlp.hashtagtreinamentos.com/python/intensivao/login"
email = "fernando@hotmail.com"
senha = "ksjncjksckdbchbdcdb"

# Acessar o site
pg.PAUSE = 0.5 # Velocidade a depender do processamento do PC
pg.press("win")
pg.write("chrome")
pg.press("enter")
pg.write(link)
pg.press("enter")

# Realizar o login
 # pyautogui.countdown(2) - Depende da velocidade em que o site abre
pg.click(x=638, y=472) # Ou pyautogui.press("tab")
pg.write(email)
pg.press("tab")
pg.write(senha)
pg.click(x=941, y=658) # Ou pyautogui.press("tab") -> pyautogui.press("enter")5.0       

# Criar variável com a planilha
tabela = pd.read_csv("produtos.csv")

# Cadastrar o produto
for linha in tabela.index:
    pg.click(x=627, y=345) # Ou pyautogui.press("tab")
    # Código do produto
    pg.write(tabela.loc[linha, "codigo"])
    pg.press("tab")
    # Marca do produto
    pg.write(tabela.loc[linha, "marca"])
    pg.press("tab")
    # Tipo do produto
    pg.write(tabela.loc[linha, "tipo"])
    pg.press("tab")
    # Categoria do produto
    pg.write(str(tabela.loc[linha, "categoria"]))
    pg.press("tab")
    # Preço unitário do produto
    pg.write(str(tabela.loc[linha, "preco_unitario"]))
    pg.press("tab")
    # Custo do produto
    pg.write(str(tabela.loc[linha, "custo"]))
    pg.press("tab")
    # OBS
    obs = tabela.loc[linha, "obs"]
    if not pd.isna(obs):
        pg.write(obs)
    pg.press("tab")
    # Enviar o produto
    pg.countdown(1)
    pg.press("enter", presses = 2, interval = 0.5)
    # Voltar para o início da tela
    pg.countdown(1)
    pg.scroll(5000)