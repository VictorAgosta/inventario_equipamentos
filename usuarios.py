from tkinter import ttk
import pandas as pd
from tkinter import *
from tkinter import messagebox
from registros import cadastro_registro

######################################### Cadastro de Usuários #########################################

# criando as variaveis de listas que receberão os dados
email_lista = []
nome_lista = []
cod_pessoa_lista = []
ramal_lista = []
setor_lista = []
status_lista = []


# função para atualizar a tabela usuarios
def atualizar_plan_pessoas():
    plan0_pessoas = pd.read_excel('Inventario_usuario.xlsx', sheet_name='usuario')

    # inserção de código automático
    ult_cod_pessoa = plan0_pessoas[plan0_pessoas.columns[0]].count()
    cont_nomes = 0
    while cont_nomes < len(nome_lista):
        cont_nomes += 1
        cod_pessoa_lista.append(f'US{(ult_cod_pessoa + cont_nomes)}')

    # criação do dicionário
    dicio_pessoas = {'cod_usuario': cod_pessoa_lista, 'nome': nome_lista,
                     'setor': setor_lista, 'email': email_lista, 'ramal': ramal_lista, 'status': status_lista}
    df_pessoas = pd.DataFrame(dicio_pessoas)

    # atualização da planilha
    plan1_pessoas = pd.concat([plan0_pessoas, df_pessoas], ignore_index=True)
    plan1_pessoas.to_excel('Inventario_usuario.xlsx', sheet_name='usuario', index=False)

    # limpar as listas para caso queira atualizar com mais valores a planilha
    cod_pessoa_lista.clear()
    nome_lista.clear()
    email_lista.clear()
    setor_lista.clear()
    ramal_lista.clear()


# função da interface para preenchimento dos dados
def cadastro_pessoa():
    # criação da interface cadastro pessoas
    app1 = Toplevel()
    app1.title('Cadastro de Pessoas')
    app1.geometry('500x500')
    app1.configure(background='Silver')

    # nome:
    lb1 = Label(app1, text='Nome:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=30, height=20, width=100)

    nome = Entry(app1)
    nome.place(x=120, y=30, height=20, width=350)

    # setor
    lb2 = Label(app1, text='Setor:', background='Silver', foreground='black', anchor=W)
    lb2.place(x=10, y=60, height=20, width=100)

    setor = Entry(app1)
    setor.place(x=120, y=60, height=20, width=350)

    # e-mail
    lb3 = Label(app1, text='E-mail:', background='Silver', foreground='black', anchor=W)
    lb3.place(x=10, y=90, height=20, width=100)

    email = Entry(app1)
    email.place(x=120, y=90, height=20, width=350)

    # ramal
    lb4 = Label(app1, text='Ramal:', background='Silver', foreground='black', anchor=W)
    lb4.place(x=10, y=120, height=20, width=100)

    ramal = Entry(app1)
    ramal.place(x=120, y=120, height=20, width=350)

    def comando_btn11():
        nome_lista.append(nome.get())
        setor_lista.append(setor.get())
        email_lista.append(email.get())
        ramal_lista.append(ramal.get())
        status_lista.append('Ativo')
        nome.delete(0, "end")
        setor.delete(0, "end")
        email.delete(0, "end")
        ramal.delete(0, "end")

    btn11 = Button(app1, text='Carregar Dados', command=comando_btn11)
    btn11.place(x=100, y=170, height=40, width=150)

    # botão para atualizar a planilha com os dados carregados
    btn12 = Button(app1, text='Atualizar Planilha', command=atualizar_plan_pessoas)
    btn12.place(x=300, y=170, height=40, width=150)

    app1.mainloop()
