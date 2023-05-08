import pandas as pd
import os
import boto3
from datetime import datetime, timedelta
import PySimpleGUI as sg
import ctypes
import pyperclip

ctypes.windll.kernel32.FreeConsole()

diretorio_final = os.getcwd() + '\\StatusBolt\\'

ACCESS_KEY = 'YOUR_ACCESS_KEY'
SECRET_KEY = 'YOUR_SECRET_KEY'
s3 = boto3.client('s3', aws_access_key_id=ACCESS_KEY, aws_secret_access_key=SECRET_KEY)

bucket_name = "validabolt"

paginator = s3.get_paginator('list_objects')
page_iterator = paginator.paginate(Bucket=bucket_name)
filtered_iterator = page_iterator.search("Contents")

lista = []

for key_data in filtered_iterator:
    lista.append(key_data)

ls = pd.DataFrame(lista)
ls.columns = ['id', 'data', '3', '4', '5', '6']
ls1 = ls.drop('3', axis=1)
ls2 = ls1.drop('4', axis=1)
ls3 = ls2.drop('5', axis=1)
ls4 = ls3.drop('6', axis=1)
ls5 = ls4.drop('data', axis=1)

i1 = ls5['id'].str.split('_', expand=True)
i2 = i1.drop(1, axis=1)
i3 = i2.drop(2, axis=1)
i3.rename(columns={0: "Bolt"}, inplace=True)

i4 = i2
i5 = i4.drop(2, axis=1)
i5.rename(columns={0: "ID"}, inplace=True)

jac = i1
jac.rename(columns={1: "firm"}, inplace=True)
jac.rename(columns={2: "freq"}, inplace=True)

col_firmware = jac["firm"]
col_frequency = jac["freq"]

oi = ls4.astype(str)
oi2 = oi['data'].str.split('+', expand=True)
oi3 = oi2.drop(1, axis=1)
oi1 = oi3[0].str.split(' ', expand=True)

extracted_col2 = i3['Bolt']
extracted_col3 = i5["ID"]

vincula = pd.read_excel(diretorio_final + "IdBolt.xlsx")
vincula1 = pd.DataFrame(vincula)

oi3[0] = pd.to_datetime(oi3[0], format='%Y-%m-%d %H:%M:%S')
oi3[0] = oi3[0] - timedelta(hours=3)
oi3[0] = pd.to_datetime(oi3[0], format='%Y-%m-%d %H:%M:%S')
oi3[0] = oi3[0].dt.strftime('%d-%m-%Y %H:%M:%S')

pipo = oi3[0].str.split(':', expand=True)
pipo1 = pipo.drop(1, axis=1)
pipo2 = pipo1.drop(2, axis=1)

oi4 = oi3[0].str.split(' ', expand=True)
oi4.rename(columns = {0: "Data", 1: "Horário"}, inplace=True)

hoje1 = datetime.today()
hoje2 = hoje1.strftime('%d/%m/%Y')
hoje3 = hoje2.replace("/", "-")

def checar_data(Data):
    if Data == hoje3:
        return 'On'
    else:
        return 'Offline'

oi4['Status'] = oi4['Data'].apply(checar_data)

oi5 = oi4.join(extracted_col2)
oi32 = oi5.join(extracted_col3)
oi34 = oi32.join(col_firmware)
oi35 = oi34.join(col_frequency)
oi6 = oi35.sort_values('Status', ascending=False)

df_f = oi6[['Bolt', 'Data', 'Horário', 'Status', 'ID', 'firm', 'freq']]

df_f = df_f.loc[df_f['ID'] != 'LANG123456']

for valor in range(len(df_f)):
    item = df_f.iloc[valor, 0]

    for id_nome in vincula1.values:
        if item == str(id_nome[0]):

            df_f.iloc[valor, 0] = id_nome[1]

headers = {'Bolt': [], 'Data': [], 'Horário': [], 'Status': [], 'ID': [], 'Firm': [], 'Freq': []}
headings = list(headers)

tablee = df_f
data = tablee.values.tolist()

sg.set_options(font=("Arial", 9))

image_restart = diretorio_final + 'buscar.png'

layout =[[sg.Text("Basta clicar nos elementos para copiar!", text_color='#FDFDFF', background_color='#434347', border_width=3)],
        [sg.Table(data, headings=headings, auto_size_columns=True, enable_events=True, enable_click_events=True, justification='center', num_rows=13, key='-CONTACT_TABLE-', background_color='#ECF0F1', text_color='#17202A', row_height=45)],
        [sg.Text("Nome da empresa: ", text_color='#17202A', background_color='#ECF0F1'),
         sg.Input(size=(18, 1), background_color='#ECF0F1', key='-INPUT-'),
         sg.ReadFormButton('.', image_filename=image_restart, button_color='#ECF0F1', image_size=(25, 25), image_subsample=2, border_width=0),
         sg.Button('Vincular Nomes / ID', button_color='#12205F', border_width=0,), sg.Button('Planilha CSV', button_color='#12205F', border_width=0), sg.Button('Abrir Diretório', button_color='#12205F', border_width=0)]]

df_f['ID'] = df_f['ID'].astype(float)

icon = diretorio_final + 'ibbx.ico'

window = sg.Window('Status Bolt', layout, auto_close=True, background_color='#ECF0F1', grab_anywhere=True, auto_close_duration=3540, finalize=True, icon=icon)

table = window['-CONTACT_TABLE-']
entry = window['-INPUT-']
entry.bind('<Return>', 'RETURN-')

df_f.to_csv(diretorio_final + "StatusBolt.csv", sep=';', header=True, index=False, encoding='latin1')

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break

    elif isinstance(event, tuple) and event[:2] == ('-CONTACT_TABLE-', '+CLICKED+'):
        row, col = position = event[2]
        if None not in position and row >= 0:
            text = data[row][col]
            pyperclip.copy(text)

    if event in ('.', '-INPUT-RETURN-'):
        text = values['-INPUT-'].lower()
        if text == '':
            continue
        row_colors = []
        for row, row_data in enumerate(data):
            if text in row_data[0].lower():
                row_colors.append((row, '#99A3A4'))
            else:
                row_colors.append((row, '#ECF0F1'))
        table.update(row_colors=row_colors)

    print(diretorio_final)
    if event == 'Abir Diretório':
        os.startfile(diretorio_final)
    if event == 'Vincular Nomes / ID':
        os.startfile(diretorio_final + 'IdBolt.xlsx')
    if event == 'Planilha CSV':
        os.startfile(diretorio_final + 'StatusBolt.csv')

window.close()
