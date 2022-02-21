from typing import Text
from PIL import Image, ImageFont, ImageDraw
import barcode
import pandas as pd
from barcode import EAN14, EAN13
from barcode.writer import ImageWriter
import os
import pyautogui
import shutil
#Localizção caixa master
Icodigo_produto = (1039, 194)
InFPL = (278, 352)
Iquantidade = (263, 650)
Icod_barras_master = (843, 541)
#Localização etiqueta embalagem
Ititulo = (22, 200)
Idescricao = (22, 255)
Icodigo_de_barras_embalagem = (50, 490)
InumCodgoBarras = (160, 675)
Icod_embalagem = (709, 490)
Iqtd_embalagem = (568, 580)
#Fontes caixa Master
caminho_fonte= "BebasNeue-Regular.ttf"
fontFPL = ImageFont.truetype(caminho_fonte, 60)
fontCod = ImageFont.truetype(caminho_fonte, 200)
fontQt = ImageFont.truetype(caminho_fonte, 200)
cor = (0,0,0)
#Fontes Embalagem
FonteTitulo = ImageFont.truetype(caminho_fonte, 50)
FonteDescricao = ImageFont.truetype(caminho_fonte, 30)
FonteCodEmbalagem = ImageFont.truetype(caminho_fonte, 90)
FonteQtdEmbalagem = ImageFont.truetype(caminho_fonte, 45)
#Abertura das planilhas
planilha = pd.read_excel("planilha.xlsx")
planilhaMatriz = pd.read_excel("Matriz.xlsx")
Pfpl = planilha.loc[0, "FPL"]
#Verificação dos codigos
y = 0
D = 0
for i, j in enumerate(planilhaMatriz["Cod"]):
    for s, m in enumerate(planilha["COD"]):
        d= +1
        if j == m:
             y = +1
    D = d
    d = 0
if y != D:
    pyautogui.alert("Precisa Atualiza a Planinha MAXTER")
else:
#Criação das pastas
    os.mkdir(Pfpl)
    Cm1 = f"{Pfpl}\\Caixa Maxter"
    Cm2 = f"{Pfpl}\\Embalagem"
    os.mkdir(Cm1)
    os.mkdir(Cm2)
    #Criação Etiquetas Embalagem
    for i, cd in enumerate(planilha["COD"]):
       #Abertura da imagem
        imagem1 = Image.open(r'mxt707.png')
        desenho1 = ImageDraw.Draw(imagem1)
        f = 0
        QtdE = planilha.loc [i, "QtdE"]
        CodigoBarrasEmbalagem = planilha.loc[i, "CodBE"]
        for s, cd2 in enumerate(planilhaMatriz["Cod"]):
            if cd2 == cd:
                #Leitura das linhas e colunas
                titulo = planilhaMatriz.loc[s, "Titulo"]
                descricao = planilhaMatriz.loc[s, "Descricao"]
                embalagem = planilhaMatriz.loc[s, "Embalagem"]
                if embalagem == ("Blister" or "blister"):
                    #Fazer copia do blister
                    shutil.copy(f"Blister\\{cd}.pdf", f"{Cm2}", follow_symlinks= True)
                else:
                    #Desenho dos textos
                    desenho1.text(Ititulo, f"{titulo}", font=FonteTitulo, fill=cor)
                    desenho1.text(Idescricao, f"{descricao}", font=FonteDescricao, fill=cor)
                    desenho1.text(Iqtd_embalagem, f"Pacote com {QtdE}", font=FonteQtdEmbalagem, fill=cor)
                    desenho1.text(Icod_embalagem,f"{cd}", font=FonteCodEmbalagem, fill=cor )
                    desenho1.text(InumCodgoBarras, f"{CodigoBarrasEmbalagem}", font=FonteDescricao, fill=cor)
                    codigo_barra_embalagem = EAN13 (f"{CodigoBarrasEmbalagem}", writer=ImageWriter())
                    codigo_barra_embalagem.save(f"cb{cd}")
                    codigo_barra2_embalagem = Image.open(f"cb{cd}.png")
                    codigo_barra3_embalagem=codigo_barra2_embalagem.crop(box=(70, 4, 454, 194 ))
                    #Salve
                    im1 = imagem1.convert('RGB')
                    im1.paste(codigo_barra3_embalagem, Icodigo_de_barras_embalagem)
                    im1.save(f"{Cm2}\\{cd}.PDF")
                    os.remove(f"cb{cd}.png")
    #Criação etiqueta caixa master
    for i, qt in enumerate(planilha["QTD"]):
        #Abertura da imagem
        imagem = Image.open(r'mxt708.png')
        desenho = ImageDraw.Draw(imagem)
        #Leitura das linhas e colunas
        Pcodigo = planilha.loc[i, "COD"]
        Pcb = planilha.loc[i, "CODB"]
        #Desenho dos textos
        desenho.text(Icodigo_produto, f"{Pcodigo}", font=fontCod, fill=cor)
        desenho.text(InFPL, Pfpl, font=fontFPL, fill=cor)
        desenho.text(Iquantidade, f"{qt}", font=fontQt, fill=cor)
        #Criação do codigo de barras
        codigo_barra = EAN14 (f"{Pcb}", writer=ImageWriter())
        codigo_barra.save(f"cb{Pcodigo}")
        codigo_barra2 = Image.open(f"cb{Pcodigo}.png")
        #Salvando imagem 
        im = imagem.convert('RGB')
        im.paste(codigo_barra2, Icod_barras_master)
        im.save(f"{Cm1}\\{Pcodigo}.PDF")
        os.remove(f"cb{Pcodigo}.png")