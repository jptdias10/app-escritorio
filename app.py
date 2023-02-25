import pandas as pd
import numpy as np
import openpyxl
from tqdm import tqdm
from typing import Dict, List
file = 'Caixa.xlsm'
workbook = openpyxl.load_workbook(file)
sheets = workbook.sheetnames
#TODO loop para ler página a página
#TODO levando em conta que há 2 modelos, faça uma função para identificar pela data se é do formato antigo ou novo
sheet = workbook['MARÇO-16'].values
pag_toda = pd.DataFrame(sheet)
pag_toda.replace({None:np.nan}, inplace=True)
