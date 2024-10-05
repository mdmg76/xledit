import pandas as pd
# import numpy as np
# import openpyxl
# import csv


source_csv = 'test_csv_new.csv'
pksz_db = "//network/path/to/file/Pack Size Database.xlsx"
source_xl = 'test_xl.xlsx'


df_csv = pd.read_csv(source_csv, sep=';', names=list(
    'abcdefghijklmnopqrstuvw'))
df_pksz = pd.read_excel(pksz_db)
df_xl = pd.read_excel(source_xl)

df_csv.drop(['v', 'w'], axis=1, inplace=True)

promoted_header = df_csv.iloc[0]
df_csv = df_csv[1:]
df_csv.columns = promoted_header

df_csv.columns.values[[1, 2, -1]] = ['Partial Per Pack', 'NDC', 'PHX']

df_csv.drop(df_csv.columns.difference(
    ['Quantity', 'Partial Per Pack', 'NDC', 'PHX']), axis=1, inplace=True)

df_csv.drop_duplicates(inplace=True)

df_csv[['Open', 'Loose']] = df_csv['Partial Per Pack'].str.split(
    ' / ', expand=True)
df_csv.drop(['Partial Per Pack'], axis=1, inplace=True)

df_csv[['Quantity', 'Open']] = df_csv[[
    'Quantity', 'Open']].apply(pd.to_numeric)

df_csv['Packs'] = df_csv['Quantity'] - df_csv['Open']

df_csv.drop(['Quantity', 'Open'], axis=1, inplace=True)
df_csv.drop(df_csv[df_csv.NDC == 'SWEEPER'].index, inplace=True)

df_csv = df_csv.merge(df_pksz, how='left', on='NDC')

df_csv.loc[df_csv['PHX_x'] != df_csv['PHX_y'], 'PHX Code'] = df_csv['PHX_y']
df_csv.loc[df_csv['PHX_x'] == df_csv['PHX_y'], 'PHX Code'] = df_csv['PHX_x']

df_csv.drop(['PHX_x', 'PHX_y'], axis=1, inplace=True)
df_csv = df_csv[['NDC', 'PHX Code', 'Loose', 'Packs', 'Pack Size', 'UOM']]

df_csv['Pack Size'] = df_csv['Pack Size'].str.split('|')
df_csv = df_csv.set_index(['NDC', 'PHX Code', 'Loose', 'Packs', 'UOM'])['Pack Size'].apply(
    pd.Series).stack().reset_index().rename(columns={0: 'Pack Size'})

df_csv = df_csv[['NDC', 'PHX Code', 'Loose', 'Packs', 'Pack Size', 'UOM']]

df_csv[['Packs', 'Pack Size', 'Loose']] = df_csv[[
    'Packs', 'Pack Size', 'Loose']].apply(pd.to_numeric)

df_csv.loc[df_csv['UOM'] == 'ea', 'QOH'] = (
    df_csv['Packs']*df_csv['Pack Size'])+df_csv['Loose']
df_csv.loc[df_csv['UOM'] != 'ea', 'QOH'] = df_csv['Packs']

df_csv.drop(df_csv.columns.difference(
    ['PHX Code', 'UOM', 'QOH']), axis=1, inplace=True)


df_csv = df_csv.groupby(['PHX Code', 'UOM'])['QOH'].sum().reset_index()

writer = pd.ExcelWriter('rowa_stock.xlsx', engine='xlsxwriter')
df_csv.to_excel(writer, sheet_name='Current Rowa Stock',
                startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets['Current Rowa Stock']

(max_row, max_col) = df_csv.shape

column_settings = []
for header in df_csv.columns:
    column_settings.append({'header': header})

worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
worksheet.set_column(0, max_col - 1, 12)

writer.save()

print(df_csv.head(500))
