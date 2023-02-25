import pandas as pd
import numpy as np
import openpyxl
from tqdm import tqdm
from typing import Dict, List

def remove_none_rows_and_cols(df):
    df = df.dropna(axis=1, how='all')
    df = df.dropna(axis=0, how='all')
    return df

def row0_to_header( dfin: pd.DataFrame)->pd.DataFrame:
    """ Get the top row and rename the columns with it
    Args:
        df (pd.DataFrame): DataFrame to replace header
    Returns:
        pd.DataFrame: DataFrame with header renamed
    """
    new_header = dfin.iloc[0] #grab the first row for the header
    dfout = dfin[1:].copy() #take the data less the header row
    dfout.columns = new_header #set the header row as the df header
    return dfout
def get_ixinit_saidas(df:pd.DataFrame):
    """Retorna os índices dos inícios das tabelas 'Entradas' e 'Saídas'

    Args:
        df (pd.DataFrame): _description_

    Returns:
        List[index]: Indícies dos inícios das tabelas 'Entradas' e 'Saídas'
    """
    if (df.iloc[:,0].str.contains('SAÍDAS') == True).any():
        ixinit_saidas = (df.iloc[:,0].str.contains('SAÍDAS') == True).idxmax()
    else:
        ixinit_saidas = None
    return ixinit_saidas

def get_entradas_saidas(df:pd.DataFrame)->List[pd.DataFrame]:
    """Separa duas tabelas com os índices de início já identificados

    Args:
        df (pd.DataFrame):
        ixinits (List[int]):

    Returns:
        List[pd.DataFrame]:
    """
    df = remove_none_rows_and_cols(df)
    ixinits = get_ixinit_saidas(df)
    print(ixinits)
    entradas = df.iloc[:ixinits-2] #TODO Função get_entradas que já retorna a tabela entradas toda tratada
    print('entradas')
    saidas = df.iloc[ixinits-2:] #TODO Função get_saidas que já retorna a tabela saidas toda tratada
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
worksheet = workbook['MARÇO-16'] #TODO Remover
pag_toda = pd.read_excel('Caixa.xlsm', sheet_name='MARÇO-16', engine='openpyxl')
pag_toda.replace({None:np.nan}, inplace=True)
#TODO se for do antigo, encapsular vvv em função
fechamento = get_fechamento(pag_toda.iloc[:,:2])
adiantamentos = get_adiantamento(pag_toda.iloc[:,3:5])
tables: List[pd.DataFrame] = get_entradas_saidas(pag_toda.iloc[:,6:12])
entradas = tables[0]
saidas = tables[1]

#TODO se for do novo, encapsular vvv em função

