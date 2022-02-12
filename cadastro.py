import datetime as dt
import time
import tkinter as tk
from tkinter import ttk
import sqlite3
import pandas as pd
from tkinter import *

#criando o banco de dados, a conexao e a tabela
#con = sqlite3.connect('bd_producao.db')

#cursor = con.cursor()

#cursor.execute(''' CREATE TABLE producao(data text, unidade text, especialidade text, profissional text, dia_semana text,
                 #turno text, status text, total_agendamento text, total_encaixe text, total_faltas text, data_criacao text)''')

#con.commit()

#con.close()

#função para cadastrar no banco de dados
def cadastrar():
    con = sqlite3.connect('bd_producao.db')
    cursor = con.cursor()
    cursor.execute("INSERT INTO producao values(:data,:unidade,:especialidade, :profissional,:dia_semana,:turno,:status,:total_agendamento,:total_encaixe,:total_faltas, :data_criacao)",
                   {'data':input_data.get(),
                    'unidade':input_unidade.get(),
                    'especialidade':input_especialidade.get(),
                    'profissional': input_profissional.get(),
                    'dia_semana':input_dia_semana.get(),
                    'turno':input_turno.get(),
                    'status':input_status.get(),
                    'total_agendamento':input_total_agendamento.get(),
                    'total_encaixe':input_total_encaixe.get(),
                    'total_faltas':input_total_faltas.get(),
                    'data_criacao':dt.datetime.now()})

    con.commit()
    con.close()
    #comandos abaixo limpam os dados dos campos apos clicar em salvar no banco
    input_especialidade.delete(0, 'end')
    input_profissional.delete(0 , 'end')
    input_status.delete(0,'end')
    input_total_agendamento.delete(0,'end')
    input_total_encaixe.delete(0,'end')
    input_total_faltas.delete(0,'end')

#função exportar

def exportaExcel():
    con = sqlite3.connect('bd_producao.db')
    cursor = con.cursor()
    cursor.execute("SELECT *, oid FROM producao")
    dados_bd = cursor.fetchall()
    dados_bd = pd.DataFrame(dados_bd, columns= ['data','unidade','especialidade', 'profissional','dia_semana','turno','status','total_agendamento','total_encaixe','total_faltas', 'data_criacao', 'id'])
    dados_bd.to_excel('planilha_producao.xlsx', index=False)
    con.commit()
    con.close()

dados_producao = []
lista_data = [1,2,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]
lista_status = ["Consulta Realizada", "Remarcado a Pedido", "Remarcado por Falta", "Remarcado por Doenca", "Remarcado Erro_Agenda"]
lista_turno =["Manha", "Tarde"]
lista_dias = ["Segunda-Feira", "Terca-Feira", "Quarta-Feira", "Quinta-Feira", "Sexta-Feira"]
lista_unidades =["Policlinica", "Centro Materno"]
lista_especialidade = ["Alergista","Angiologia","Cardiologia","Cirurgia_Cabeça_Pescoço","Cirurgia_Geral","Cirurgia_Plastico",
                      "Cirurgia_Vascular","Clinica_Medica","Dermatologia","Endocrinologia","Fonoaudiologia","Gastroenterologia",
                      "Geriatria","Ginecologia","Hematologia","Homeopatia","Infectologia","Pericia_Medica","Mastologia","Medicina_do_Trabalho",
                       "Nefrologia","Neurocirurgia","Neurologia","Neuropediatria","Nutricionista","Obstetricia","Oftalmologia",
                       "Ortopedia","Otorrino","Pediatria","Pequenas_cirurgias","Pneumologia","Proctologia","Psicologia",
                       "Psiquiatria","Reumatologia","Urologia","USG"]

# Janela principal
janela = Tk()
#janela.geometry('850x250')
janela.configure(bg='grey')

#dimensões da janela:
largura = 1130
altura = 500
#capturando dimensões da janela do pc
largura_screen = janela.winfo_screenwidth()
altura_screen = janela.winfo_screenheight()

#posição da janela:
posx = largura_screen/2 - largura/2
posy = altura_screen/2 - altura/2

janela.geometry("%dx%d+%d+%d" % (largura, altura, posx, posy))

#Criação da função
def inserir_codigo():
    data = input_data.get()
    unidade = input_unidade.get()
    especialidade = input_especialidade.get()
    profissional = input_profissional.get()
    dia_semana = input_dia_semana.get()
    turno = input_turno.get()
    status = input_status.get()
    total_agendamento = input_total_agendamento.get()
    total_encaixe = input_total_encaixe.get()
    total_faltas = input_total_faltas.get()

    data_criacao = dt.datetime.now()
    data_criacao = data_criacao.strftime("%d/%m/%Y %H:%M")
    codigo = len(dados_producao)+1
    codigo_str = "COD-{}".format(codigo)
    dados_producao.append((codigo_str, data, unidade, especialidade, profissional, dia_semana, turno, status, total_agendamento, total_encaixe,
    total_faltas, data_criacao))

#Título da Janela,
janela.title('   @BuziosCloudTecnologia                                                                          Ferramenta Básica para cadastro de Produção.                                             Desenvolvida por @Otto Vieira    ')

data = tk.Label(text="Data:")
data.grid(row=0, column=0,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_data = ttk.Combobox(values=lista_data)
input_data.grid(row=0, column=2, padx = 10, pady=10, sticky='nswe', columnspan = 2)

dia_semana = tk.Label(text="Dia Semana:")
dia_semana.grid(row=0, column=4,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_dia_semana = ttk.Combobox(values=lista_dias)
input_dia_semana.grid(row=0, column=6, padx = 10, pady=10, sticky='nswe', columnspan = 2)

turno = tk.Label(text="Turno:")
turno.grid(row=0, column=8,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_turno = ttk.Combobox(values=lista_turno)
input_turno.grid(row=0, column=10, padx = 10, pady=10, sticky='nswe', columnspan = 2)

unidade = tk.Label(text="Unidade de saúde:")
unidade.grid(row=0, column=12, padx=10, pady=10, sticky='nswe', columnspan=2)
input_unidade = ttk.Combobox(values=lista_unidades)
input_unidade.grid(row=0, column=14, padx=10, pady=10, sticky='nswe', columnspan=2)

especialidade = tk.Label(text="Especialidade:")
especialidade.grid(row=1, column=0,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_especialidade = ttk.Combobox(values=lista_especialidade)
input_especialidade.grid(row=1, column=2, padx = 10, pady=10, sticky='nswe', columnspan = 2)

profissional = tk.Label(text="Nome Profissional")
profissional.grid(row=1, column=4, padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_profissional = tk.Entry()
input_profissional.grid(row=1, column=6,padx = 10, pady=10, sticky='nswe', columnspan =5 )

status = tk.Label(text="Status:")
status.grid(row=1, column=11,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_status = ttk.Combobox(values=lista_status)
input_status.grid(row=1, column=13, padx = 10, pady=10, sticky='nswe', columnspan = 4)

total_agendamento = tk.Label(text="Total_Agendamento")
total_agendamento.grid(row=2, column=0,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_total_agendamento = tk.Entry()
input_total_agendamento.grid(row=2, column=2,padx = 10, pady=10, sticky='nswe', columnspan =2 )

total_encaixe = tk.Label(text="Total_Encaixe")
total_encaixe.grid(row=2, column=4,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_total_encaixe = tk.Entry()
input_total_encaixe.grid(row=2, column=6,padx = 10, pady=10, sticky='nswe', columnspan =2 )

total_faltas = tk.Label(text="Total de Faltas")
total_faltas.grid(row=2, column=8,padx = 10, pady=10, sticky='nswe', columnspan =2 )
input_total_faltas = tk.Entry()
input_total_faltas.grid(row=2, column=10,padx = 10, pady=10, sticky='nswe', columnspan =2 )


btn_salvar = tk.Button(text="Salvar", command=cadastrar)
btn_salvar.grid(row=13,column=0,padx = 10, pady=10,sticky='nswe', columnspan =3)

btn_exportar = tk.Button(text="Exportar Excel", command=exportaExcel)
btn_exportar.grid(row=13, column=4,padx = 10, pady=10,sticky='nswe', columnspan =3)

janela.mainloop()


