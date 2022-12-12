import pandas as pd
from tkinter import filedialog
import datetime

data = datetime.date.today()

#Abrir MC
criado = filedialog.askopenfilename()
with open(criado, 'rb') as file:
    textfromfile = file.read()
df = pd.read_excel(criado)

#IMPORTANDO E LIMPANDO CRIADO MC
df = df.drop([0, 1, 2], axis=0)
df = df.drop(['Unnamed: 3','Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8'], axis=1)
df.columns = ('Cliente','Endereço','Região','Horario','Codigo')
df = df[df['Região'].isin(['02', '05'])]
df = df.astype('string')
df['nova'] = df['Endereço'].str.extract(r'([A-Z, ""]{3})')
df['cod'] = df['Região'].str.strip() + df['Codigo'].str.strip() + df['nova']
df = df.drop(['Cliente', 'Endereço', 'Codigo', 'nova', 'Horario'], axis=1)
df['Sequência'] = range(len(df))
df.columns = ('Região', 'Código do Cliente', 'Sequência')

#MONTANDO PLANILHA ROUTE
df1 = pd.read_excel('BASE AGF.xlsx')
df2 = pd.merge(df, df1, on='Código do Cliente', how='left')
df2 = df2[['Código do Cliente', 'Nome do Cliente', 'Rua', 'Bairro', 'Número', 'Complemento', 'Município', 'CEP', 'Região', 'Sequência']]

df2.to_excel(f'Rota AGF {data.day}-{data.month}.xlsx', index=False)
