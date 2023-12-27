import tkinter as tk
from tkinter import Entry, ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
from datetime import datetime


requisicao = requests.get("https://economia.awesomeapi.com.br/json/all")
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())
#lista_moedas = [(sigla, nome) for sigla, nome in dicionario_moedas.items()]

def pegar_cotacao():
    moeda_escolhida = combobox_selecionarMoeda.get()
    data_escolhida = calendario_data.get()
    dia = data_escolhida[:2]
    mes = data_escolhida[3:5]
    ano = data_escolhida[-4:]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda_escolhida}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    resposta = requests.get(link)
    resposta = resposta.json()
    cotacao = resposta[0]["bid"]
    lable_respostacotacao["text"] = f"A Cotação da moeda {moeda_escolhida} no dia {data_escolhida} é: R$ {cotacao}"
    

def selecionar_arquivo():
    caminho = askopenfilename(title="Selecione o arquivo das moedas")
    var_caminhoArquivo.set(caminho) #Salvamos esse caminho dentro da variavel criada no Tkinter
    if caminho:
        lable_arquivoSelecionado["text"] = f"Arquivo selecionado: {caminho}"


def atualizar_cotacoes():
    try:    
        df = pd.read_excel(var_caminhoArquivo.get())
        moedas = df.iloc[:,0] #[linhas:colunas] = Todas as linhas, coluna 0 (A)

        periodo = int(quant_dias.get())

        for cada_moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{cada_moeda}-BRL/{periodo}"
            cotacoes = requests.get(link)
            cotacoes = cotacoes.json()

            for cada_dia in cotacoes:
                timestamp_cotacao = int(cada_dia["timestamp"])
                bid = float(cada_dia["bid"])
                data = datetime.fromtimestamp(timestamp_cotacao)
                data = data.strftime("%d/%m/%Y")
            
                if data not in df:
                    df[data] = ""

                #selecionando uma celula especifica do DF - Linha x Coluna
                #linha da coluna A == cada_moeda e a coluna referente a data
                df.loc[df.iloc[:,0]==cada_moeda, data] = bid
        
        df.to_excel("Teste.xlsx") #onde salva o arquivo
        lable_atualizarCotacoes['text'] = "Arquivo Atualizado com Sucesso"

    except:
        lable_atualizarCotacoes['text'] = "Erro ao ler o arquivo. Não está no padrão."


janela = tk.Tk()

#Nome da Janela
janela.title("Sistema de Cotação de Moedas")

#Titulo da parte 1  do Sistema
lable_CotacaoMoeda = tk.Label(text="Cotação de 1 moeda específica", borderwidth=2, relief="solid")
lable_CotacaoMoeda.grid(row=0, column=0, padx=10, pady=10, sticky="nswe", columnspan=3)

#Selecionar Moeda
lable_selecionarMoeda = tk.Label(text="Selecione a moeda da Cotação:", anchor="e")
lable_selecionarMoeda.grid(row=1, column=0, padx=10, pady=10, sticky="nswe", columnspan=2)

#ComboBOx para selecionar moeda
combobox_selecionarMoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarMoeda.grid(row=1, column=2,padx=10, pady=10, sticky="nswe")

#Selecionar dia da cotação
lable_dataCotacao = tk.Label(text="Selecione o dia da Cotação:", anchor="e")
lable_dataCotacao.grid(row=2, column=0, padx=10, pady=10, sticky="nswe", columnspan=2)

#Calendario para Escolha da data
calendario_data = DateEntry(year=2023, locale="pt_br")
calendario_data.grid(row=2, column=2, padx=10, pady=10, sticky="nsew")

#lable Texto Resposta da Cotação
lable_respostacotacao = tk.Label(text="")
lable_respostacotacao.grid(row=3, column=0, padx=10, pady=10, sticky="nswe", columnspan=2)

#botão Pegar Cotação
botao_pegarCotacao = tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegarCotacao.grid(row=3, column=2, padx=10, pady=10,  sticky="nswe")


#Titulo da parte 2 do Sistema
lable_CotacaoVariasMoedas = tk.Label(text="Cotação de Multiplas moeda específica", borderwidth=2, relief="solid")
lable_CotacaoVariasMoedas.grid(row=4, column=0, padx=10, pady=10, sticky="nswe", columnspan=3)

#lable Selecionar arquivo
lable_selecionarArquivo = tk.Label(text="Selecione um arquivo em Excel com as moedas na Coluna A")
lable_selecionarArquivo.grid(row=5, column=0, padx=10, pady=10, sticky="nswe", columnspan=2)

#botão Selecionar Arquivo
var_caminhoArquivo = tk.StringVar() #Irá armazenar o caminho completo que o usuário escolher - usado dentro da função "selecionar_arquivo"
botao_selecionarArquivo = tk.Button(text="Click para Selecionar", command=selecionar_arquivo)
botao_selecionarArquivo.grid(row=5, column=2, padx=10, pady=10,  sticky="nswe")

#lable arquivo Selecionado
lable_arquivoSelecionado = tk.Label(text="Nenhum arquivo selecionado", anchor="e")
lable_arquivoSelecionado.grid(row=6, column=0, padx=10, pady=10, sticky="nswe", columnspan=3)

#lable Escolha do Período:
lable_quantDias = tk.Label(text="Cotar os ultimos X dias: ", anchor="e")
lable_quantDias.grid(row=7, column=0, padx=10, pady=10, sticky="nswe")

#Entdada da quantidade de dias desejado
quant_dias = tk.Entry()
quant_dias.grid(row=7, column=1, padx=10, pady=10, sticky="nsew")


#botão Atualizar Cotações
botao_atualizarCotacoes = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizarCotacoes.grid(row=9, column=0, padx=10, pady=10,  sticky="nswe")

#lable Texto Resposta de Atualizar Cotaçoes
lable_atualizarCotacoes = tk.Label(text="")
lable_atualizarCotacoes.grid(row=9, column=1, padx=10, pady=10, sticky="nswe", columnspan=2)

#botão fechar Programa
botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10,  sticky="nswe")

janela.mainloop()