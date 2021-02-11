#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
import os
import sys


filename = 'QBEntry.xlsx'
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)
file_path = os.path.join(application_path, filename)
        
xl = pd.ExcelFile(file_path)
styles = pd.read_excel(xl, 'Styles').dropna(how='all')
colors = pd.read_excel(xl, 'Colors').dropna(how='all').set_index('Shoe Style')
descriptions = pd.read_excel(xl, 'Descriptions').dropna(how='all').set_index('Shoe Style')
prices = pd.read_excel(xl, 'Prices').dropna(how='all').set_index('Shoe Style')
sizes = pd.read_excel(xl, 'Sizes').dropna(how='all')




df_for_qb = pd.DataFrame(columns = ['Item', 'Description', 'Price', '', ' '])
columns = df_for_qb.columns

for name in styles['Name'].dropna():
    color_list = list(colors.loc[name].dropna())
    size_list = [size[0] for size in sizes.values]
    description = descriptions.loc[name].values[0]
    price = prices.loc[name].values[0]
    for color in color_list:
        for size in size_list:
            item = '{name} {color} LS-{size}'.format(name=name, color=color, size=size)
            new_row = [item, description, price, '', 'yes']
            df = pd.DataFrame([new_row], columns = columns)
            df_for_qb = df_for_qb.append(df)
            
exit_filename = 'Excel File for Quickbooks Entry.xlsx'
exit_path = os.path.join(application_path, exit_filename)
df_for_qb.to_excel(exit_path, index=False)