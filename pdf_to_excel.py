import pdftables_api
import pandas as pd
from os import path
import warnings
import numpy as np

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

#pdf para excel
if path.isfile('Report.pdf'):
    print("Arquivo já existente")
else:
    c = pdftables_api.Client('rpkmac9nc77d')
    c.xlsx('Report.pdf', 'Report')
    print("Arquivo criado")


#excel 
arq_excel = "Report.xlsx"
dataframes = pd.read_excel(arq_excel,sheet_name=None)
dados_combinados = pd.concat(dataframes.values(),ignore_index=True)
df= pd.DataFrame(dados_combinados)
remove = [" ̈","1/1","Itens Rateio","MP","Máquina","Tp. Serviço"]
for frase in remove:
    df.replace(frase,np.nan,regex=True,inplace=True)
df=df.drop(colums=df.columns[df.columns.get_loc("C"):df.columns.get_loc("E"+1)])
df=df.dropna(how="all")
df.apply(lambda x: x.dropna().reset_index(drop=True), axis=1)

df= df.apply(lambda x: pd.Series(x.dropna().values),axis=1)

dados_combinados = df
dados_combinados.to_excel("Report_att.xlsx",index=False,engine="openpyxl")


print(dados_combinados)
