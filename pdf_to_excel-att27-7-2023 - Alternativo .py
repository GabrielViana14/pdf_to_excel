import aspose.pdf as ap
import pandas as pd
import os
from os import system
import numpy as np

input_pdf = "Report.pdf"
output_pdf ="Report_Alt.xlsx"
# Open PDF document
document = ap.Document(input_pdf)
save_option = ap.ExcelSaveOptions()
document.save(output_pdf, save_option)
header_new = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17]

arq_excel = "Report_Alt.xlsx"
dataframes = pd.read_excel(arq_excel,sheet_name=None)
dados_combinados = pd.concat(dataframes.values(),ignore_index=True)
df= pd.DataFrame(dados_combinados)

df.to_excel(output_pdf,index=False,header=header_new)


remove = [" ̈","1/1","No desenho","Itens Rateio","MP","Máquina","Tp. Serviço",'MSM CALDEIRARIA - www.msmcaldeiraria.com.br','CNPJ.: 06183455000176 IE.: 581.067.044.113']
for frase in remove:
    df.replace(frase,np.nan,regex=True,inplace=True)
coldel= [1,2,3,4,6,11,13,14,15,16]
df=df.drop(df.columns[coldel],axis=1)
df = df.reset_index(drop=True)
df = df.dropna(how='all')
#df = df[~df['Unnamed: 6'].str.contains("Total de Horas :", case=False)]



df.to_excel("Report_att_alt.xlsx",index=False,engine="openpyxl")
nomearq = "Report_att_alt.xlsx"

if os.path.exists(nomearq):
    os.system(f"start {nomearq}")
else:
    print("Arquivo não encontrado.")
    print(df)



