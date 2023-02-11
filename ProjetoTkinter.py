import tkinter as tk
from tkinter import ttk
import numpy as np
from tkcalendar import DateEntry
import requests
from tkinter.filedialog import askopenfilename
import pandas as pd
from datetime import datetime

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


# cotação de uma moeda específica

def pegar_cotacao():
    moeda = combobox_moeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    mensagem3["text"] = f"A cotação da {moeda} no dia {data_cotacao} foi de: R$ {valor_moeda}"


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione um arquivo em Excel para abrir')
    var_caminho_arquivo.set(caminho_arquivo)
    if caminho_arquivo:
        mensagem5["text"] = f"Aquivo selecionado: {caminho_arquivo}"

def atualizar_cotacao():
    df = pd.read_excel(var_caminho_arquivo.get())
    moedas = df.iloc[:, 0]
    data_inicial = calendario_moeda2.get()
    data_final = calendario_moeda3.get()

    ano_inicial = data_inicial[-4:]
    mes_inicial = data_inicial[3:5]
    dia_inicial = data_inicial[:2]

    ano_final = data_final[-4:]
    mes_final = data_final[3:5]
    dia_final = data_final[:2]

    d_inicial = datetime.strptime(data_inicial, '%d/%m/%Y')
    d_final = datetime.strptime(data_final, '%d/%m/%Y')
    
    num_dias = abs((d_final-d_inicial).days)+1

    for moeda in moedas:
        link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/{num_dias}?" \
               f"start_date={ano_inicial}{mes_inicial}{dia_inicial}&" \
               f"end_date={ano_final}{mes_final}{dia_final}"

        requisicao_moeda = requests.get(link)
        cotacoes = requisicao_moeda.json()
        print(link)
        for cotacao in cotacoes:
            timestamp = int(cotacao['timestamp'])
            bid = float(cotacao['bid'])
            data = datetime.fromtimestamp(timestamp)
            data = data.strftime('%d/%m/%Y')
            if data not in df:
                df[data] = np.nan

            df.loc[df.iloc[:, 0] == moeda, data] = bid
    df.to_excel("Teste.xlsx")
    mensagem9["text"] = "Arquivo atualizado com sucesso"

janela = tk.Tk()

janela.title("Ferramenta de Cotação de Moedas")

janela.rowconfigure([0,10], weight=1)
janela.columnconfigure([0, 2], weight=1)

titulo1 = tk.Label(text="Cotação de 1 moeda específica", bg='black', borderwidth=2, relief='solid')
titulo1.grid(row=0, column=0, padx=10, pady=10, columnspan=3, sticky='NSEW')

mensagem1 = tk.Label(text="Selecione a moeda que deseja consultar:", bg='darkblue')
mensagem1.grid(row=1, column=0, padx=10, pady=10, columnspan=2, sticky='NSEW')

combobox_moeda = ttk.Combobox(values=lista_moedas)
combobox_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='NSEW')

mensagem2 = tk.Label(text="Selecione o dia que deseja pegar a cotação:", bg='darkblue')
mensagem2.grid(row=2, column=0, padx=10, pady=10, columnspan=2, sticky='NSEW')

calendario_moeda = DateEntry(year=2023, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='NSEW')

mensagem3 = tk.Label(text="", bg='darkblue')
mensagem3.grid(row=3, column=0, padx=10, pady=10, columnspan=2, sticky='NSEW')

botao1= tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao1.grid(row=3, column=2, padx=10, pady=10,sticky='NSEW')

#cotação de várias moedas

titulo2 = tk.Label(text="Cotação de Multiplas Moedas", bg='black', borderwidth=2, relief='solid')
titulo2.grid(row=4, column=0, padx=10, pady=10, columnspan=3, sticky='NSEW')

mensagem4 = tk.Label(text="Selecione um arquivo em Excel com as Moedas na Coluna A", bg='darkblue')
mensagem4.grid(row=5, column=0, padx=10, pady=10, columnspan=2, sticky='NSEW')

var_caminho_arquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text='Clique para Selecionar', command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10,sticky='NSEW')

mensagem5 = tk.Label(text="Arquivo Selecionado:", bg='darkblue')
mensagem5.grid(row=6, column=0, padx=10, pady=10, columnspan=3, sticky='NSEW')

mensagem6 = tk.Label(text="Data Inicial:", bg='darkblue')
mensagem6.grid(row=7, column=0, padx=10, pady=10, columnspan=1, sticky='NSEW')

calendario_moeda2 = DateEntry(year=2023, locale='pt_br')
calendario_moeda2.grid(row=7, column=2, padx=10, pady=10, sticky='NSEW')

mensagem7 = tk.Label(text="Data Final:", bg='darkblue')
mensagem7.grid(row=8, column=0, padx=10, pady=10, columnspan=1, sticky='NSEW')

calendario_moeda3 = DateEntry(year=2023, locale='pt_br')
calendario_moeda3.grid(row=8, column=2, padx=10, pady=10, sticky='NSEW')

mensagem8 = tk.Label(text="Arquivo Selecionado:", bg='darkblue')
mensagem8.grid(row=9, column=0, padx=10, pady=10, columnspan=2, sticky='NSEW')

botao2= tk.Button(text='Atualizar Cotações', command=atualizar_cotacao)
botao2.grid(row=9, column=0, padx=10, pady=10,sticky='NSEW')

mensagem9 = tk.Label(text="", bg='darkblue')
mensagem9.grid(row=9, column=1, padx=10, pady=10, columnspan=2, sticky='NSEW')

botao3= tk.Button(text='Fechar', command=janela.quit)
botao3.grid(row=11, column=2, padx=10, pady=10,sticky='NSEW')

janela.mainloop()