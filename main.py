import pandas as pd
import os
import boto3
from datetime import datetime, timedelta
import PySimpleGUI as sg
import ctypes
import pyperclip

ctypes.windll.kernel32.FreeConsole()

caminho_diretorio = os.getcwd() + '\\StatusBolt\\'

CHAVE_ACESSO = 'YOUR_ACCESS_KEY'
CHAVE_SECRETA = 'YOUR_SECRET_KEY'
cliente_s3 = boto3.client('s3', aws_access_key_id=CHAVE_ACESSO, aws_secret_access_key=CHAVE_SECRETA)

nome_bucket = "validabolt"

paginador = cliente_s3.get_paginator('list_objects')
iterador_paginas = paginador.paginate(Bucket=nome_bucket)
iterador_filtrado = iterador_paginas.search("Contents")

lista_objetos = []

for dados_chave in iterador_filtrado:
    lista_objetos.append(dados_chave)

frame_dados = pd.DataFrame(lista_objetos)
frame_dados.columns = ['id', 'dados', '3', '4', '5', '6']
frame_sem_3 = frame_dados.drop('3', axis=1)
frame_sem_4 = frame_sem_3.drop('4', axis=1)
frame_sem_5 = frame_sem_4.drop('5', axis=1)
frame_sem_6 = frame_sem_5.drop('6', axis=1)
frame_sem_dados = frame_sem_6.drop('dados', axis=1)

split_id = frame_sem_dados['id'].str.split('_', expand=True)
frame_id_parcial = split_id.drop(1, axis=1)
frame_id_final = frame_id_parcial.drop(2, axis=1)
frame_id_final.rename(columns={0: "Parafuso"}, inplace=True)

frame_id_completo = split_id
frame_id_sem_2 = frame_id_completo.drop(2, axis=1)
frame_id_sem_2.rename(columns={0: "ID"}, inplace=True)

frame_firmware = split_id
frame_firmware.rename(columns={1: "firmware", 2: "frequencia"}, inplace=True)

coluna_firmware = frame_firmware["firmware"]
coluna_frequencia = frame_firmware["frequencia"]

frame_str = frame_sem_5.astype(str)
frame_split_dados = frame_str['dados'].str.split('+', expand=True)
frame_split_dados = frame_split_dados.drop(1, axis=1)
frame_data_hora = frame_split_dados[0].str.split(' ', expand=True)

col_parafuso = frame_id_final['Parafuso']
col_id = frame_id_sem_2["ID"]

vinculos = pd.read_excel(caminho_diretorio + "IdBolt.xlsx")
frame_vinculos = pd.DataFrame(vinculos)

frame_split_dados[0] = pd.to_datetime(frame_split_dados[0], format='%Y-%m-%d %H:%M:%S')
frame_split_dados[0] = frame_split_dados[0] - timedelta(hours=3)
frame_split_dados[0] = pd.to_datetime(frame_split_dados[0], format='%Y-%m-%d %H:%M:%S')
frame_split_dados[0] = frame_split_dados[0].dt.strftime('%d-%m-%Y %H:%M:%S')

frame_split_hora = frame_split_dados[0].str.split(':', expand=True)
frame_hora_sem_1 = frame_split_hora.drop(1, axis=1)
frame_hora_sem_2 = frame_hora_sem_1.drop(2, axis=1)

frame_data_horario = frame_split_dados[0].str.split(' ', expand=True)
frame_data_horario.rename(columns={0: "Data", 1: "Hora"}, inplace=True)

hoje_formatado = datetime.today().strftime('%d/%m/%Y').replace("/", "-")

def verificar_status(data):
    if data == hoje_formatado:
        return 'Online'
    else:
        return 'Offline'

frame_data_horario['Status'] = frame_data_horario['Data'].apply(verificar_status)

frame_resultado = frame_data_horario.join(col_parafuso)
frame_resultado = frame_resultado.join(col_id)
frame_resultado = frame_resultado.join(coluna_firmware)
frame_resultado = frame_resultado.join(coluna_frequencia)
frame_resultado = frame_resultado.sort_values('Status', ascending=False)

resultado_final = frame_resultado[['Parafuso', 'Data', 'Hora', 'Status', 'ID', 'firmware', 'frequencia']]

resultado_final = resultado_final.loc[resultado_final['ID'] != 'LANG123456']

for indice in range(len(resultado_final)):
    parafuso = resultado_final.iloc[indice, 0]

    for id_nome in frame_vinculos.values:
        if parafuso == str(id_nome[0]):
            resultado_final.iloc[indice, 0] = id_nome[1]

cabecalhos = {'Parafuso': [], 'Data': [], 'Hora': [], 'Status': [], 'ID': [], 'Firmware': [], 'Frequência': []}
cabecalhos_lista = list(cabecalhos)

tabela = resultado_final
dados_tabela = tabela.values.tolist()

sg.set_options(font=("Arial", 9))

imagem_atualizar = caminho_diretorio + 'buscar.png'

layout = [
    [sg.Text("Basta clicar nos elementos para copiar!", text_color='#FDFDFF', background_color='#434347', border_width=3)],
    [sg.Table(dados_tabela, headings=cabecalhos_lista, auto_size_columns=True, enable_events=True, enable_click_events=True, justification='center', num_rows=13, key='-TABELA_CONTATOS-', background_color='#ECF0F1', text_color='#17202A', row_height=45)],
    [sg.Text("Nome da empresa: ", text_color='#17202A', background_color='#ECF0F1'),
     sg.Input(size=(18, 1), background_color='#ECF0F1', key='-ENTRADA-'),
     sg.ReadFormButton('.', image_filename=imagem_atualizar, button_color='#ECF0F1', image_size=(25, 25), image_subsample=2, border_width=0),
     sg.Button('Vincular Nomes / ID', button_color='#12205F', border_width=0), sg.Button('Planilha CSV', button_color='#12205F', border_width=0), sg.Button('Abrir Diretório', button_color='#12205F', border_width=0)]
]

resultado_final['ID'] = resultado_final['ID'].astype(float)

icone = caminho_diretorio + 'ibbx.ico'

janela = sg.Window('Status Parafuso', layout, auto_close=True, background_color='#ECF0F1', grab_anywhere=True, auto_close_duration=3540, finalize=True, icon=icone)

tabela = janela['-TABELA_CONTATOS-']
entrada = janela['-ENTRADA-']
entrada.bind('<Return>', 'RETURN-')

resultado_final.to_csv(caminho_diretorio + "StatusBolt.csv", sep=';', header=True, index=False, encoding='latin1')

while True:
    evento, valores = janela.read()
    if evento == sg.WINDOW_CLOSED:
        break

    elif isinstance(evento, tuple) and evento[:2] == ('-TABELA_CONTATOS-', '+CLICKED+'):
        linha, coluna = posicao = evento[2]
        if None not in posicao and linha >= 0:
            texto_copiado = dados_tabela[linha][coluna]
            pyperclip.copy(texto_copiado)

    if evento in ('.', '-ENTRADA-RETURN-'):
        texto = valores['-ENTRADA-'].lower()
        if texto == '':
            continue
        cores_linhas = []
        for linha, dados_linha in enumerate(dados_tabela):
            if texto in dados_linha[0].lower():
                cores_linhas.append((linha, '#99A3A4'))
            else:
                cores_linhas.append((linha, '#ECF0F1'))
        tabela.update(row_colors=cores_linhas)

    print(caminho_diretorio)
    if evento == 'Abrir Diretório':
        os.startfile(caminho_diretorio)
    if evento == 'Vincular Nomes / ID':
        os.startfile(caminho_diretorio + 'IdBolt.xlsx')
    if evento == 'Planilha CSV':
        os.startfile(caminho_diretorio + 'StatusBolt.csv')

janela.close()
