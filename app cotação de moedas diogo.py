import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import requests
import json
from datetime import datetime

## janela
janela = tk.Tk()


## título
janela.title('Ferramenta de Cotação de Moedas App Diogo')


## redimensionamento automático da janela
janela.rowconfigure([0,1,2,3,4,5,6,7,8,9,10], weight=1)
janela.columnconfigure([0, 1, 2], weight=1)


## placeholders, functions (back-end)
requisicao_api = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_api = requisicao_api.json()
lista_moedas = list(dicionario_api.keys())

def pegar_cotacao():
    moeda_escolhida = listasuspensa1.get()
    data_escolhida = calendario1.get()
    data_escolhida = data_escolhida.split('/')
    ano = data_escolhida[2]
    mes = data_escolhida[1]
    dia = data_escolhida[0]
    requisicao_api2 = requests.get(f'https://economia.awesomeapi.com.br/json/daily/{moeda_escolhida}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}')
    dicionario_api2 = requisicao_api2.json() # uma lista com um dicionário
    cotacao = dicionario_api2[0]['bid']
    mensagem1['text'] = f"Cotação da moeda {moeda_escolhida} no dia {dia}/{mes}/{ano}: R$ {cotacao}"


def selecionar_arquivo():
    arquivo = askopenfilename(title='Selecione um arquivo Excel')
    variavel_armazenamento_arquivo.set(arquivo)
    if arquivo:
        mensagem2['text'] = f"Arquivo selecionado: {arquivo}"
        mensagem2['foreground'] = '#0000FF'


def atualizar_cotacoes():
    try:
        df = pd.read_excel(variavel_armazenamento_arquivo.get())
        lista_moedas_selecionadas = df.iloc[:,0]
        qtde_dias = input1.get()

        for moeda in lista_moedas_selecionadas:
            requisicao_api3 = requests.get(f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/{qtde_dias}')
            dicionario_api3 = requisicao_api3.json()  # uma lista com vários dicionários

            for dicionario in dicionario_api3:
                cotacao = float(dicionario['bid'])
                dia_da_cotacao = int(dicionario['timestamp'])
                dia_da_cotacao = datetime.fromtimestamp(dia_da_cotacao)
                dia_da_cotacao = dia_da_cotacao.strftime('%d/%m/%Y')

                if dia_da_cotacao not in df:
                    df[dia_da_cotacao] = ''

                df.loc[df.iloc[:,0]==moeda, dia_da_cotacao] = cotacao

        print(df)
        df.to_excel('Cotações das Moedas Final.xlsx', index=False)
        mensagem3['text'] = 'Arquivo de Cotações Atualizado'
        mensagem3['foreground'] = '#0000FF'

    except:
        mensagem3['text'] = 'Selecione um arquivo Excel no formato correto e/ou digite um número correto para a quantidade de dias'
        mensagem3['foreground'] = '#FF0000'


def finalizar():
    janela.quit()


## edição geral da janela (front-end)
texto1 = tk.Label(text='Cotação de 1 moeda', background='#5b5b5b', foreground='#F1C232', borderwidth=3, relief='solid')
texto1.grid(row=0, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

texto2 = tk.Label(text='Selecionar a moeda:', anchor='e')
texto2.grid(row=1, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

listasuspensa1 = ttk.Combobox(janela, values=lista_moedas)
listasuspensa1.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

texto3 = tk.Label(text='Selecionar o dia que deseja pegar a cotação:', anchor='e')
texto3.grid(row=2, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

calendario1 = DateEntry(year=2023, locale='pt_br')
calendario1.grid(row=2, column=2, padx=10, pady=10, sticky='nsew')

mensagem1 = tk.Label(text='', foreground='#0000FF')
mensagem1.grid(row=3, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

botao1 = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao1.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

texto4 = tk.Label(text='Cotação de várias moedas', background='#5b5b5b', foreground='#F1C232', borderwidth=3, relief='solid')
texto4.grid(row=4, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

texto5 = tk.Label(text='Selecionar arquivo Excel com as moedas desejadas na coluna A:')
texto5.grid(row=5, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

variavel_armazenamento_arquivo = tk.StringVar() # variável onde, na function, será armazenado o caminho do arquivo excel selecionado
botao2 = tk.Button(text='Selecionar Arquivo', command=selecionar_arquivo)
botao2.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

mensagem2 = tk.Label(text='Nenhum arquivo selecionado', anchor='e', foreground='#FF0000')
mensagem2.grid(row=6, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

texto6 = tk.Label(text='Digite a quantidade de dias desejada de cotações:', anchor='e')
texto6.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')

input1 = tk.Entry()
input1.grid(row=7, column=1, padx=10, pady=10, sticky='nsew')

botao3 = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao3.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

mensagem3 = tk.Label(text='', anchor='w')
mensagem3.grid(row=9, column=1, padx=10, pady=10, sticky='nsew', columnspan=2)

botao4 = tk.Button(text='Finalizar', command=finalizar)
botao4.grid(row=10, column=1, padx=10, pady=10)

## visualizar janela
janela.mainloop()