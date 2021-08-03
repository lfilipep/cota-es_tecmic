#! python3
'''
cotacoes ficheiro prazo_de_entrega

Dá as melhores as melhores cotações inseridas no ficheiro. Desde que cumpram os requisitos.
O MOQ tem de ser superior à quantidade ou o fornecedor ser a Tecmic e o prazo de entrega do componente
tem de ser inferior ao prazo de entrega estipulado.

ficheiro: Ficheiro Excel com os componentes a analisar. Tem de seguir o template estabelecido.
prazo_de_entrega: Número de dias máximo para entrega dos componentes.
'''

import os
import sys
from datetime import datetime, timedelta, date
from pprint import pprint
from numpy import isnan
import pandas as pd
from pandas.core.dtypes.missing import notnull


prazo_entrega = int(sys.argv[2]) 
CURR_DIR = os.getcwd()
cotações_original_file = os.path.join(CURR_DIR, sys.argv[1])
csv_data = {
    'Referencia': [],
    'Designacao': [],
    'Observacoes': []
}


data_xls = pd.read_excel(cotações_original_file, 'Folha1', dtype=str)
data_xls.to_csv('csvfile.csv', encoding='utf-8', index=False)

data = pd.read_csv('csvfile.csv')
data_best_price = pd.DataFrame()
header = data.columns

elements = data['Referencia'].drop_duplicates()

ind_count = data['Preço'].count()
data['divider'] = data['QT'] / data['MOQ']


for ind in range(ind_count):
    if data['divider'].iloc[ind] > 1:
        product = data['divider'].iloc[ind]
        product_int = data['divider'].iloc[ind].astype(int)
        if product % product_int > 0:
            data.at[ind, 'divider'] = product_int + 1 # update cell values without errors
        else:
            data.at[ind, 'divider'] = product_int
    else:
        data.at[ind, 'divider'] = 1.0    

data['Total_preço'] = data['Preço']* data['MOQ'] * data['divider']
data['Qt a encomendar'] = 0 

for ind in range(ind_count):
    if data['Preço'].iloc[ind] > 0:
        data.at[ind, 'Qt a encomendar'] = data['MOQ'].iloc[ind] * data['divider'].iloc[ind] 
        
        


del data['divider']

for e in elements:
    components = pd.DataFrame()
    # data_best_price = data_best_price.append(data[data['Referencia'] == e].sort_values(by='Total_preço', ascending=True).head(1), ignore_index=True)
    components = components.append(data[data['Referencia'] == e], ignore_index=True) # Group data by component
    referencia = components['Referencia'].iloc[0]
    designacao = components['Designacao'].iloc[0]
    quantidade = components['QT'].iloc[0]
    # components_moq = components[(components['MOQ'] >= components['QT']) | ((components['Fornecedor'] == 'Tecmic') & (components['MOQ'] < components['QT']))] 
    if not components.empty:
        components_deadline_date = components[components['Prazo (dias)'] <= prazo_entrega]
        if not components_deadline_date.empty:
            componentes_price_zero = components_deadline_date[components_deadline_date['Total_preço'] == 0]
            data_best_price = data_best_price.append(componentes_price_zero, ignore_index=True)
            componentes_price_greater_zero = components_deadline_date[components_deadline_date['Total_preço'] > 0].sort_values(by='Total_preço', ascending=True)
            data_best_price = data_best_price.append(componentes_price_greater_zero.head(1), ignore_index=True)
        else:
            csv_data['Referencia'].append(referencia)
            csv_data['Designacao'].append(designacao)
            csv_data['Observacoes'].append(f"Prazo de entrega superior a {prazo_entrega} dias")
    else:
        csv_data['Referencia'].append(referencia)
        csv_data['Designacao'].append(designacao)
        csv_data['Observacoes'].append(f"MOQ inferior a {quantidade}")


data_best_price.to_excel(os.path.join(CURR_DIR, "melhores cotacoes.xlsx"), header=True, index=False)

no_requisites_data = pd.DataFrame(csv_data)
no_requisites_data.to_excel(os.path.join(CURR_DIR, "componentes fora dos requisitos.xlsx"), header=True, index=False)

