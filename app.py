import pandas as pd
import numpy as np
import openpyxl
from tqdm import tqdm
from typing import Dict, List
file = 'Caixa.xlsm'
workbook = openpyxl.load_workbook(file)
sheets = workbook.sheetnames
