#! python3

import os
import sys
from datetime import datetime, timedelta, date
from pprint import pprint
from numpy import isnan
import pandas as pd
from pandas.core.dtypes.missing import notnull



cotações_original_file = "L:/Python Projects/Tecmic Project/cotações.xlsx"
prazo_entrega = 90
# prazo_entrega = sys.argv[2]

CURR_DIR = os.getcwd()
print(CURR_DIR)
# cotações_original_file = os.path.join(SAVE_DIR, sys.argv[1]')

# Create Analyze Dasboard File
# analyze_dashboard = open(os.path.join(SAVE_DIR, f'Analyze Summary {today}.csv'), 'w')
# analyze_dashboard.write("Num_Frota;ID;Total_Mens;Num_POWER_ONs;Num_POWER_OFFs;Num_IGN_ONs;Num_IGN_OFFs;WithoutGPS;WithoutGPS(%);DelayMens;MensSameMinute;Service\n")
# analyze_dashboard.close()

data_xls = pd.read_excel(cotações_original_file, 'Folha1', dtype=str)
data_xls.to_csv('csvfile.csv', encoding='utf-8', index=False)

data = pd.read_csv('csvfile.csv')
data_best_price = pd.DataFrame()
header = data.columns

elements = data['Referencia'].drop_duplicates()

data['Total_preço'] = data['Preço'] * data['MOQ']

with open("componentes_sem_requisitos.txt", "a") as report:

    for e in elements:
        components = pd.DataFrame()
        # data_best_price = data_best_price.append(data[data['Referencia'] == e].sort_values(by='Total_preço', ascending=True).head(1), ignore_index=True)
        components = components.append(data[data['Referencia'] == e], ignore_index=True) # Group data by component
        referencia = components['Referencia'].iloc[0]
        designacao = components['Designacao'].iloc[0]
        quantidade = components['QT'].iloc[0]
        components_moq = components[(components['MOQ'] > components['QT']) | ((components['Fornecedor'] == 'Tecmic') & (components['MOQ'] < components['QT']))] 
        if not components_moq.empty:
            components_deadline_date = components_moq[components_moq['Prazo (dias)'] < prazo_entrega]
            if not components_deadline_date.empty:
                componentes_price_zero = components_deadline_date[components_deadline_date['Total_preço'] == 0]
                data_best_price = data_best_price.append(componentes_price_zero, ignore_index=True)
                componentes_price_greater_zero = components_deadline_date[components_deadline_date['Total_preço'] > 0].sort_values(by='Total_preço', ascending=True)
                data_best_price = data_best_price.append(componentes_price_greater_zero.head(1), ignore_index=True)
            else:
                report.write(f"{referencia} - {designacao} tem prazo de entrega mais elevado que {prazo_entrega} dias\n")
        else:
            report.write(f"{referencia} - {designacao} tem um MOQ inferior a {quantidade}\n")

    data_best_price.to_excel("L:/Python Projects/Tecmic Project/cotações best price.xlsx", header=True, index=False)


# with open('csvfile.csv', "r") as csvfile:
#    csvreader = csv.reader(csvfile, delimiter=";")
#    next(csvreader)

#    for row in list(csvreader):
#       for r in row:
#           print(r.split(',')[0])
