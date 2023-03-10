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

def verifica_indices(indices_inicio, indices_fim):
    if len(indices_fim) > len(indices_inicio):
        for i in range(len(indices_inicio)):
            if indices_inicio[i] >= indices_fim[i]:
                print(f"Erro: o índice de início {indices_inicio[i]} é maior ou igual ao índice de fim {indices_fim[i]} na posição {i}")
                indices_fim.pop(i)
                return [indices_inicio, indices_fim]
        print("As listas têm tamanhos diferentes, mas os índices estão em ordem crescente")
        return [indices_inicio, indices_fim]
    else:
        print("As listas têm o mesmo tamanho ou a lista de índices de início é maior que a lista de índices de fim")
        return [indices_inicio, indices_fim]

def fill_date(df:pd.DataFrame)->pd.DataFrame:
    """Fill NaN with min mode of column

    Args:
        df (pd.DataFrame): _description_

    Returns:
        _type_: _description_
    """
    return df['DATA'].fillna(df['DATA'].mode().min())

def get_ixinit_saidas(df:pd.DataFrame):
#     """Retorna os índices dos inícios das tabelas 'Entradas' e 'Saídas'

#     Args:
#         df (pd.DataFrame): _description_

#     Returns:
#         List[index]: Indícies dos inícios das tabelas 'Entradas' e 'Saídas'
#     """
    indices_dt = get_data_indices(df)
    ixinit_entradas = indices_dt[0]
    ixinit_saidas = indices_dt[1]
    if (df.iloc[:,0].str.contains('TOTAL') == True).any():
        ixend_entradas = (df.iloc[:,0].str.contains('TOTAL') == True).idxmax()
    else:
        ixend_entradas = None
    return [ixinit_entradas, ixend_entradas, ixinit_saidas]

def get_entradas(df:pd.DataFrame)->pd.DataFrame:
    entradas = row0_to_header(df)
    entradas = entradas.drop(entradas[entradas['DATA'] == 'TOTAL'].index)
    entradas = entradas.dropna(subset=['VALOR'])
    entradas['DATA'] = fill_date(entradas)
    entradas = remove_none_rows_and_cols(entradas)
    return entradas

def get_saidas(df:pd.DataFrame):
    saidas = row0_to_header(df)
    saidas = saidas.dropna(subset=['VALOR'])
    saidas['DATA'] = fill_date(saidas)
    saidas = saidas[['DATA', 'VALOR', 'MOTIVO']]
    saidas = remove_none_rows_and_cols(saidas)
    return saidas

def get_entradas_saidas(df:pd.DataFrame)->List[pd.DataFrame]:
    """Separa duas tabelas com os índices de início já identificados

    Args:
        df (pd.DataFrame):
        ixinits (List[int]):

    Returns:
        List[pd.DataFrame]:
    """
    ixinits = get_ixinit_saidas(df)
    entradas = get_entradas(df.iloc[ixinits[0]:ixinits[1]])
    saidas = get_saidas(df.iloc[ixinits[2]:])
    return [entradas, saidas]

def get_total_indices(df:pd.DataFrame):
    return df.loc[df.iloc[:, 0] == 'TOTAL'].index

def get_data_indices(df:pd.DataFrame):
    return df.loc[df.iloc[:, 0] == 'DATA'].index

def get_adiantamento(df:pd.DataFrame)->pd.DataFrame:
    #TODO Adicionar coluna de motivo caso nao tenha
    li_dfs = []
    adiantamento = remove_none_rows_and_cols(df)
    adiantamento = row0_to_header(adiantamento)
    adiantamento.reset_index(inplace=True)
    adiantamento.drop('index', axis=1, inplace=True)
    indices = verifica_indices(get_data_indices(adiantamento),
                               get_total_indices(adiantamento))
    indices_inicio = indices[0]
    indices_fim = indices[1]
    for i in range(len(indices_fim)):
        if i == 0:
            df_aux = adiantamento.iloc[indices_inicio[i]:indices_fim[i]]
            pessoa = df_aux.columns[0]
            df_aux = row0_to_header(df_aux)
        else:
            df_aux = adiantamento.iloc[indices_inicio[i]-1:indices_fim[i]]
            pessoa = df_aux.iloc[0,0]
            df_aux = row0_to_header(df_aux)
            df_aux = row0_to_header(df_aux)
        df_aux['RESPONSAVEL'] = pessoa
        li_dfs.append(df_aux)
    adiantamento = pd.concat(li_dfs)
    return adiantamento

file = 'Caixa.xlsm'
workbook = openpyxl.load_workbook(file)
sheets = workbook.sheetnames
entradas:List[pd.DataFrame] = []
saidas:List[pd.DataFrame] = []
adiantamentos:List[pd.DataFrame] = []

#TODO levando em conta que há 3 modelos(pré março17, pre fev-21, atual)
for sheet in sheets:
    print(sheet)
    pag_toda = pd.read_excel('Caixa.xlsm', sheet_name=sheet, engine='openpyxl')
    pag_toda.replace({None:np.nan}, inplace=True)

    #TODO se for do antigo, encapsular vvv em função
    # fechamento = get_fechamento(pag_toda.iloc[:,:2])
    adiantamentos.append(get_adiantamento(pag_toda.iloc[:,3:5]))
    tables: List[pd.DataFrame] = get_entradas_saidas(pag_toda.iloc[:,6:12])
    entradas.append(tables[0])
    saidas.append(tables[1])
    print('OK')
print('Rodou tudo')
#TODO se for do novo, encapsular vvv em função

