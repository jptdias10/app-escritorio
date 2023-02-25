import pandas as pd
import numpy as np
import openpyxl
from tqdm import tqdm
from typing import Dict, List

def remove_none_rows_and_cols(df):
    df = df.dropna(axis=1, how='all')
    df = df.dropna(axis=0, how='all')
    return df

def get_ixinit_entradas_saidas(df:pd.DataFrame):
    """Retorna os índices dos inícios das tabelas 'Entradas' e 'Saídas'

    Args:
        df (pd.DataFrame): _description_

    Returns:
        List[index]: Indícies dos inícios das tabelas 'Entradas' e 'Saídas'
    """
    if (df[6].str.contains('ENTRADAS') == True).any():
        ixinit_entrada = (df[6].str.contains('ENTRADAS') == True).idxmax()
    else:
        ixinit_entrada = None
    if (df[6].str.contains('SAÍDAS') == True).any():
        ixinit_saidas = (df[6].str.contains('SAÍDAS') == True).idxmax()
    else:
        ixinit_saidas = None
    return [ixinit_entrada,ixinit_saidas]

def get_entradas_saidas(df:pd.DataFrame)->List[pd.DataFrame]:
    """Separa duas tabelas com os índices de início já identificados

    Args:
        df (pd.DataFrame):
        ixinits (List[int]):

    Returns:
        List[pd.DataFrame]:
    """
    df = remove_none_rows_and_cols(df)
    ixinits = get_ixinit_entradas_saidas(df)
    entradas = df.iloc[ixinits[0]:ixinits[1]-2] #TODO Função get_entradas que já retorna a tabela entradas toda tratada
    saidas = df.iloc[ixinits[1]-2:] #TODO Função get_saidas que já retorna a tabela saidas toda tratada
    return [entradas, saidas]

def get_fechamento(df):
    #TODO Função get_fechamento que já retorna a tabela fechamento toda tratada
    return remove_none_rows_and_cols(df)

def get_adiantamento(df):
    #TODO Função get_adiantamento que já retorna a tabela adiantamento toda tratada
    return remove_none_rows_and_cols(df)

file = 'Caixa.xlsm'
workbook = openpyxl.load_workbook(file)
sheets = workbook.sheetnames
#TODO loop para ler página a página
#TODO levando em conta que há 2 modelos, faça uma função para identificar pela data se é do formato antigo ou novo
sheet = workbook['MARÇO-16'].values
pag_toda = pd.DataFrame(sheet)
pag_toda.replace({None:np.nan}, inplace=True)
#TODO se for do antigo, encapsular vvv em função
fechamento = get_fechamento(pag_toda[[0,1]])
adiantamentos = get_adiantamento(pag_toda[[3,4]])
tables: List[pd.DataFrame] = get_entradas_saidas(pag_toda[[6,7,8,9,10,11]])
entradas = tables[0]
saidas = tables[1]

#TODO se for do novo, encapsular vvv em função

