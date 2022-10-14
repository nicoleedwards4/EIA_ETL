# -*- coding: utf-8 -*-
"""
Created on Sun Oct 17 19:46:45 2021

@author: nedwards1
"""
import pandas as pd
import numpy as np

# Update with most recent year of data available
end_year = 2020

# start_year = 2000

# column names to unpivot around
column_names = ['Data Year', 'Utility Number', 'Utility Name', 'Part',
                'Service Type', 'Data Type\nO = Observed\nI = Imputed',
                'State', 'Ownership', 'BA Code', 'Sector',
                'Thousand Dollars', 'Megawatthours', 'Count', 'T']
# create array for compiled dataset
main_df = pd.DataFrame(columns=column_names)

sector_types = ['COMMERCIAL', 'INDUSTRIAL', 'NON-RESIDENTIAL', 'OTHER',
                'RESIDENTIAL', 'TOTAL', 'TRANSPORTATION']

save_path = ("//fs109/rpa/Rates & Forecasting Group"
             f"/Data and Analytics/EIA/Sales_Ult_Cust"
             f"_{end_year}_Long.csv")

# save_path = ("//fs109/rpa/Rates & Forecasting Group"
#              f"/Data and Analytics/EIA/Sales_Ult_Cust_{start_year}"
#              f"_{end_year}_Long.csv")

# for y in range(start_year, end_year+1):
file_path = ("//fs109/rpa/Rates & Forecasting Group"
             f"/Data and Analytics/EIA/f861{end_year}"
             f"/Sales_Ult_Cust_{end_year}.xlsx")
data = pd.read_excel(file_path, header=None)
df = data.to_numpy()
newdf = df[2:]
headers = df[:1, :]
subheaders = df[2:3, :]
pddf = data[2:]
for sector in sector_types:
    if sector in headers:
        # finds column index of where sector starts
        result = np.where(headers == sector)
        class_ind = result[1]
        firstsector = np.where(subheaders == 'Thousand Dollars')
        ind_fsector = firstsector[1]
        fsind = ind_fsector[0]
        sects_ind = int(fsind)
        idvar_ind = int(class_ind)
        idvars = subheaders[0, :idvar_ind]
        valvarend = idvar_ind + 3
        valvars = subheaders[0, idvar_ind:valvarend]
        ndf = pddf.rename(columns=pddf.iloc[0])
        new_df = ndf.iloc[1:, np.r_[0:sects_ind, idvar_ind:valvarend]]
        new_df['Sector'] = sector
        main_df = main_df.append(new_df.reset_index())

final_columns = {'BA Code': 'Ba Code', 'Thousand Dollars':
                 'Revenues (Thousand Dollars)', 'Megawatthours':
                     'Sales (Megawatthours)', 'Count': 'Customers (Count)',
                     'Data Type\nO = Observed\nI = Imputed': 'Data Type'}

main_df = main_df.rename(columns=final_columns)
main_df = main_df[~main_df['Data Year'].str.contains('To calculate', na=False)]
main_df = main_df.iloc[:, :-2]
main_df.to_csv(save_path, index=False)
print('file updated')
