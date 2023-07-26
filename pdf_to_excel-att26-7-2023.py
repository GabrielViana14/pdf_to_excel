import pdftables_api
import pandas as pd
import os
from os import system
import warnings
import numpy as np
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

#pdf para excel
if os.path.isfile('Report.xlsx'):
    print("Arquivo já existente")
    resp = input("Deseja sobreescrever o mesmo?")
    if resp.lower() in ["s", "sim"]:
        c = pdftables_api.Client('0h5a1mj2w18n')
        c.xlsx('Report.pdf', 'Report')   
    else:
        print("Arquivo criado")

#excel 
arq_excel = "Report.xlsx"
dataframes = pd.read_excel(arq_excel,sheet_name=None)
dados_combinados = pd.concat(dataframes.values(),ignore_index=True)
df= pd.DataFrame(dados_combinados)
remove = [" ̈","1/1","No desenho","Itens Rateio","MP","Máquina","Tp. Serviço",'MSM CALDEIRARIA - www.msmcaldeiraria.com.br','CNPJ.: 06183455000176 IE.: 581.067.044.113']
for frase in remove:
    df.replace(frase,np.nan,regex=True,inplace=True)
coldel= [1,2,5]
df=df.drop(df.columns[coldel],axis=1)
df=df.dropna(how="all")
df.apply(lambda x: x.dropna().reset_index(drop=True), axis=1)

if 'Apelido:' in df.values:
    print("Valor existe")
    

df= df.apply(lambda x: pd.Series(x.dropna().values),axis=1)

idx_apelido = df[df[1] == 'Apelido:'].index[0]
idx_producao = df[df[5] == 'PRODUÇÃO'].index[0]
df.loc[idx_apelido:idx_producao,1:5] = ''
df.reset_index(drop=True, inplace=True)

while True:
    if 'Apelido:' in df[1].values and 'PRODUÇÃO' in df[5].values:
        idx_apelido = df.index[df[1] == 'Apelido:']
        idx_producao = df.index[df[5] == 'PRODUÇÃO']

        if len(idx_apelido) > 0 and len(idx_producao) > 0:
            idx_apelido = idx_apelido[0]
            idx_producao = idx_producao[0]

            df.loc[idx_apelido:idx_producao, 1:5] = ''
            df.reset_index(drop=True, inplace=True)
        else:
            break
    else:
        break

indices_funcionarios = df.index[df[0].str.startswith('Funcionário:')].tolist()
for i, indice in enumerate(indices_funcionarios):
    nome_funcionario = df.loc[indice, 0].split("Funcionário:")[1].strip()
    df.loc[indice+1:, 6] = nome_funcionario

dados_combinados = df
dados_combinados.to_excel("Report_att.xlsx",index=False,engine="openpyxl")


print(dados_combinados)
print("Report_att.xslx foi atualizado")
nomearq = "Report_att.xlsx"
if os.path.exists(nomearq):
    os.system(f'start excel "{nomearq}"')
else:
    print("Arquivo não encontrado.")
