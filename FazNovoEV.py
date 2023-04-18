from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.shared import Cm
from python_docx_replace import docx_replace
from datetime import datetime
import pandas as pd
import numpy as np


def move_table_after(table, paragraph):
    tbl, p = table._tbl, paragraph._p
    p.addnext(tbl)


# Essa função recebe uma string (texto) com valores separados por um separador (separador)
# e retorna um vetor com com os valores separados. Obs.: O último caracter da string precisa
# ser o separador.
def SeparaValores(texto, separador):
    texto_sep = []
    palavra = ""
    for caracter in texto:
        if caracter == separador:
            texto_sep.append(palavra)
            palavra = ""
        else:
            palavra = palavra + caracter

    return texto_sep


# Função que escreve o texto com a quantidade de inversores e suas respectivas potências
def EscreveTextoInv(quant_sep, pot_sep):
    if len(quant_sep) == 1:
        if quant_sep[0] == "1":
            texto_inv = "foi considerado 1 inversor de " + pot_sep[0] + " kW"
        else:
            texto_inv = "foram considerados " + quant_sep[0] + " inversores de " + pot_sep[0] + " kW"
    else:
        texto_inv = "foram considerados " + quant_sep[0] + " inversores de " + pot_sep[0] + " kW"
        for k in range(1, len(quant_sep), 1):
            if k == (len(quant_sep) - 1):
                texto_inv = texto_inv + " e " + quant_sep[k] + " inversores de " + pot_sep[
                    k] + " kW"  # escrevo inversores antes de todos?
            else:
                texto_inv = texto_inv + ", " + quant_sep[k] + " inversores de " + pot_sep[k] + " kW"
    return texto_inv


# Recebe um número em formato de texto e elimina os caracteres após a vírgula
def RoundEspecial(num_texto):
    antes_virgula = 1
    novo_texto = ""
    for k in range(len(num_texto)):
        if num_texto[k] == ",":
            antes_virgula = 0
        if antes_virgula:
            novo_texto = novo_texto + num_texto[k]
    return novo_texto


tabela = pd.read_excel("EV_SFV_INFO.xlsx", sheet_name="InfoRel")

documento = Document("EV_SFV_MODELO.docx")

for linha in tabela.index:

    titulo = str(tabela.loc[linha, "Título"])
    cliente = str(tabela.loc[linha, "Nome do Cliente"])
    endereco = str(tabela.loc[linha, "Endereço"])
    softusado = str(tabela.loc[linha, "Nome do Software utilizado"])
    notelhadocli = str(tabela.loc[linha, "Local da Instalação"])
    areadisp = str(tabela.loc[linha, "AreaDisp"])
    areadisp2 = str(tabela.loc[linha, "AreaDisp2"])
    localusi = str(tabela.loc[linha, "Local da Usina"])
    potusi = str(tabela.loc[linha, "Potência da Usina (kWp)"])
    porconaprox = str(tabela.loc[linha, "Porcentagem aproximada"])
    potmod = str(tabela.loc[linha, "Potência do Módulo (Wp)"])
    quantmod = str(tabela.loc[linha, "Quantidade de módulos"])
    modextra = str(tabela.loc[linha, "Quantidade de módulos sobressalentes"])
    marcaauto = str(tabela.loc[linha, "Marca da Automação"])
    textoestrutura = str(tabela.loc[linha, "Texto Estrutura"])
    tipoconex = str(tabela.loc[linha, "Tipo de Conexão"])
    tempmanu = str(tabela.loc[linha, "Tempo de Manutenção"])
    textosub = str(tabela.loc[linha, "Texto Subestação"])
    distri = str(tabela.loc[linha, "Distribuidora de energia do cliente"])
    datatarifa = str(tabela.loc[linha, "Data Tarifa"])
    grupocliente = str(tabela.loc[linha, "Grupo do Cliente"])
    mediager = str(tabela.loc[linha, "MEDIAGER"])
    periodocons = str(tabela.loc[linha, "Período de Consumo Analisado"])
    co2evi = str(tabela.loc[linha, "CO2 evitado"])
    arvplan = str(tabela.loc[linha, "Árvores plantadas"])
    prazoexe = str(tabela.loc[linha, "Prazo de Execução"])
    textoinje = str(tabela.loc[linha, "Texto injeção"])
    podeinje = str(tabela.loc[linha, "Texto acumular créditos"])

    potusi = RoundEspecial(potusi)
    quantmod = RoundEspecial(quantmod)
    mediager = RoundEspecial(mediager)
    co2evi = RoundEspecial(co2evi)
    arvplan = RoundEspecial(arvplan)

    quantidade = str(tabela.loc[0, "Quantidade de Inversores"])
    potencia = str(tabela.loc[0, "Potência dos Inversores"])
    quant_sep = SeparaValores(quantidade, ";")
    pot_sep = SeparaValores(potencia, ";")
    texto_inv = EscreveTextoInv(quant_sep, pot_sep)

    if quant_sep[0] == "1":
        textinv = "do Inversor"
        textinv2 = "Quantidade"
        textfig5 = "Inversor considerado"
    else:
        textinv = "dos Inversores"
        textinv2 = "Quantidade Inversores"
        textfig5 = "Inversores considerados"

    if len(pot_sep) == 1:
        textinv5 = "1 inversor"
        textinv6 = "o inversor restante permanece"
    else:
        textinv5 = "1 inversor de cada potência"
        textinv6 = "os inversores restantes permanecem"

    my_dict = {
        "NOMECLIENTETITULO": titulo,
        "NOMECLIENTE": cliente,
        "ENDEREÇO": endereco,
        "SOFTWAREUSADO": softusado,
        "NOTELHADOCLIENTE": notelhadocli,
        "AREADISP": areadisp,
        "LOCALUSINA": localusi,
        "POTUSINA": potusi,
        "AREADISP2": areadisp2,
        "PORCENCONSUMOAPROX": porconaprox,
        "POTMOD": potmod,
        "QUANTMOD": quantmod,
        "MODEXTRA": modextra,
        "TIPQUANTINV": texto_inv,
        "TEXTINV": textinv,
        "TEXTINV2": textinv2,
        # "TEXTINV3": textinv3,
        # "TEXTINV4": textinv4,
        "TEXTINV5": textinv5,
        "TEXTFIG5": textfig5,
        "TEXTINV6": textinv6,
        "MARCAAUTO": marcaauto,
        "TEXTOESTRUTURA": textoestrutura,
        "TIPOCONEXAO": tipoconex,
        "TEMPOMANUTENCAO": tempmanu,
        "TEXTOSE": textosub,
        "DISTRIBUIDORA": distri,
        "DATATARIFA": datatarifa,
        "GRUPOCLIENTE": grupocliente,
        "PERIODOCONS": periodocons,
        "MEDIAGER": mediager,
        "CO2EVI": co2evi,
        "ARVPLAN": arvplan,
        "PRAZOEXE": prazoexe,
        "TEXTOINJE": textoinje,
        "PODEINJE": podeinje,
    }

    docx_replace(documento, **my_dict)

    documento.save(f"EV_SFV_{cliente}.docx")
