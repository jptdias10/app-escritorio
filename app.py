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
def fill_date(df:pd.DataFrame)->pd.DataFrame:
    """Fill NaN with min mode of column

    Args:
        df (pd.DataFrame): _description_

    Returns:
        _type_: _description_
    """
    return df['DATA'].fillna(df['DATA'].mode().min())

def get_entradas(df:pd.DataFrame)->pd.DataFrame:
    #TODO Função get_entradas que já retorna a tabela entradas toda tratada
    entradas = row0_to_header(df)
    entradas = entradas.drop(entradas[entradas['DATA'] == 'TOTAL'].index)
    entradas['DATA'] = fill_date(entradas)
    return entradas

def get_saidas(df:pd.DataFrame):
    #TODO Função get_saidas que já retorna a tabela saidas toda tratada
    saidas = row0_to_header(df)
    saidas = saidas.dropna(subset=['VALOR'])
    saidas['DATA'] = fill_date(saidas)
    saidas = saidas[['DATA', 'VALOR', 'MOTIVO']]
    return saidas

def get_ixinit_saidas(df:pd.DataFrame):
    """Retorna os índices dos inícios das tabelas 'Entradas' e 'Saídas'

    Args:
        df (pd.DataFrame): _description_

    Returns:
        List[index]: Indícies dos inícios das tabelas 'Entradas' e 'Saídas'
    """
    if (df.iloc[:,0].str.contains('TOTAL') == True).any():
        ixend_entradas = (df.iloc[:,0].str.contains('TOTAL') == True).idxmax()
    else:
        ixend_entradas = None
    if (df.iloc[:,0].str.contains('SAÍDAS') == True).any():
        ixinit_saidas = (df.iloc[:,0].str.contains('SAÍDAS') == True).idxmax()
    else:
        ixinit_saidas = None
    return [ixend_entradas, ixinit_saidas]

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

    entradas = get_entradas(df.iloc[:ixinits[0]-1])

    #TODO Função get_saidas que já retorna a tabela saidas toda tratada
    saidas = get_saidas(df.iloc[ixinits[1]-1:])


    return [entradas, saidas]

def get_fechamento(df):
    #TODO Função get_fechamento que já retorna a tabela fechamento toda tratada
    return remove_none_rows_and_cols(df)

def get_adiantamento(df):
    #TODO Função get_adiantamento que já retorna a tabela adiantamento toda tratada
    #TODO Adicionar coluna de motivo caso nao tenha
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
print('BOOOORA')

#TODO se for do novo, encapsular vvv em função

