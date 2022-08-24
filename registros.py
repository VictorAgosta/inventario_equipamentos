######################################### Cadastro Registros #########################################
from tkinter import ttk
import pandas as pd
from tkinter import *
from tkinter import messagebox
from datetime import date
import tkinter.font as tk_font
from mensagens import msg_erro_cadastro_apag, msg_erro_cadastro, msg_erro_apagar

# criando as variaveis de listas que receberão os dados
cod_equip_registro_lista = []
nome_equip_registro_lista = []
cod_pessoa_registro_lista = []
nome_pessoa_registro_lista = []
local_registro_lista = []
data_registro_lista = []
cod_chip_registro_lista = []


def atualizar_plan_registro():
    plan0_reg = pd.read_excel('Inventario_registros.xlsx', sheet_name='registros')
    plan0_equip = pd.read_excel('Inventario_equipamento.xlsx', sheet_name='equipamento')

    # criação do dicionário
    dicio_reg = {'cod_equipamento': cod_equip_registro_lista, 'equipamento': nome_equip_registro_lista,
                 'cod_usuario': cod_pessoa_registro_lista, 'usuario': nome_pessoa_registro_lista,
                 'local': local_registro_lista, 'data_registro': data_registro_lista,
                 'cod_chip': cod_chip_registro_lista}
    df_reg = pd.DataFrame(dicio_reg)

    # atualização da planilha
    plan1_reg = pd.concat([plan0_reg, df_reg], ignore_index=True)
    plan1_reg.to_excel('Inventario_registros.xlsx', sheet_name='registros', index=False)

    for codigo in cod_equip_registro_lista:
        detectar_equip_ind = plan0_equip[plan0_equip['cod_equipamento'] == codigo].index
        plan0_equip['status'].values[detectar_equip_ind] = 'Em Uso'

    plan0_equip.to_excel('Inventario_equipamento.xlsx', sheet_name='equipamento', index=False)

    cod_equip_registro_lista.clear()
    nome_equip_registro_lista.clear()
    cod_pessoa_registro_lista.clear()
    nome_pessoa_registro_lista.clear()
    local_registro_lista.clear()
    data_registro_lista.clear()
    cod_chip_registro_lista.clear()


def cadastro_registro():
    plan0_equip = pd.read_excel('Inventario_equipamento.xlsx', sheet_name='equipamento')
    plan0_pessoas = pd.read_excel('Inventario_usuario.xlsx', sheet_name='usuario')

    # criação da interface cadastro equipamentos
    app6 = Tk()
    app6.title('Cadastro Registros')
    app6.geometry('500x500')
    app6.configure(background='Silver')
    app6.resizable(width=False, height=False)

    font_titulo = tk_font.Font(family="Lucida Grande", size=15)

    # Titulo tela

    lb0 = Label(app6, text='Registro de Entrega dos Equipamentos:', background='Silver',
                foreground='black', font=font_titulo)

    lb0.place(x=0, y=10, height=30, width=500)

    # codigo equipamento:
    lb1 = Label(app6, text='Código Equipamento:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=60, height=20, width=130)

    cod_equip_reg = Entry(app6)
    cod_equip_reg.place(x=150, y=60, height=20, width=200)

    # comando para exibir nome do equipamento atrelado ao codigo
    def comando_btn63():
        codigo_equip = cod_equip_reg.get().upper()
        detectar_equip_ind = plan0_equip[plan0_equip['cod_equipamento'] == codigo_equip].index.tolist()
        if not detectar_equip_ind:
            return nome_equip_reg.configure(text='Equipamento não cadastrado')
        nome_equip = plan0_equip.iloc[detectar_equip_ind[0]]['equipamento']
        nome_equip_reg.configure(text=nome_equip, background='silver')

    # Botão para exibir nome do equipamento atrelado ao codigo
    btn63 = Button(app6, text='Exibir equip', command=comando_btn63)
    btn63.place(x=370, y=60, height=20, width=100)

    # nome equipamento:
    lb2 = Label(app6, text='Nome Equipamento:', background='Silver', foreground='black', anchor=E)
    lb2.place(x=10, y=90, height=20, width=130)

    nome_equip_reg = Label(app6, text=' ', background='Silver', foreground='black', anchor=W)
    nome_equip_reg.place(x=150, y=90, height=20, width=350)

    # codigo usuário
    lb3 = Label(app6, text='Código Usuário:', background='Silver', foreground='black', anchor=E)
    lb3.place(x=10, y=120, height=20, width=130)

    cod_pessoa_reg = Entry(app6)
    cod_pessoa_reg.place(x=150, y=120, height=20, width=200)

    # comando para exibir nome do funcionário atrelado ao codigo
    def comando_btn62():
        codigo_pessoa = cod_pessoa_reg.get().upper()
        detectar_pessoa_ind = plan0_pessoas[plan0_pessoas['cod_usuario'] == codigo_pessoa].index.tolist()
        if not detectar_pessoa_ind:
            return nome_pessoa_reg.configure(text='Pessoa não encontrada')
        nome_pessoa = plan0_pessoas.iloc[detectar_pessoa_ind[0]]['nome']
        nome_pessoa_reg.configure(text=nome_pessoa, background='silver')

    # Botão para exibir nome do funcionário atrelado ao codigo
    btn62 = Button(app6, text='Exibir Usuario', command=comando_btn62)
    btn62.place(x=370, y=120, height=20, width=100)

    # Nome usuário
    lb4 = Label(app6, text='Nome Usuário:', background='Silver', foreground='black', anchor=E)
    lb4.place(x=10, y=150, height=20, width=130)

    nome_pessoa_reg = Label(app6, text=' ', background='Silver', foreground='black', anchor=W)
    nome_pessoa_reg.place(x=150, y=150, height=20, width=350)

    # local
    lb5 = Label(app6, text='Setor Equipamento:', background='Silver', foreground='black', anchor=E)

    lb5.place(x=10, y=180, height=20, width=130)

    local_reg = ttk.Combobox(app6, values=['Desenvolvimento', 'Segurança', 'Comercial', 'Comercial Externo',
                                           'Credenciamento', 'Financeiro', 'Gestão de Pessoas', 'Contas Médicas',
                                           'Recepção'], state="readonly")
    local_reg.place(x=150, y=180, height=20, width=320)

    def comando_btn61():
        # Definição da data da inserção dos dados modelo BR
        today = date.today()
        hoje = str(today)
        ano = hoje[:4]
        mes = hoje[5:7]
        dia = hoje[8:]

        codigo_equip = cod_equip_reg.get().upper()
        detectar_equip_ind = plan0_equip[plan0_equip['cod_equipamento'] == codigo_equip].index.tolist()
        if not detectar_equip_ind:
            return nome_equip_reg.configure(text='Equipamento não cadastrado', background='red')
        nome_equip_reg.configure(text='', background='silver')

        codigo_pessoa = cod_pessoa_reg.get().upper()
        detectar_pessoa_ind = plan0_pessoas[plan0_pessoas['cod_usuario'] == codigo_pessoa].index.tolist()
        if not detectar_pessoa_ind:
            return nome_pessoa_reg.configure(text='Pessoa não encontrada', background='red')
        nome_pessoa_reg.configure(text='', background='silver')

        cod_equip_registro_lista.append(cod_equip_reg.get().upper())
        cod_pessoa_registro_lista.append(cod_pessoa_reg.get().upper())
        local_registro_lista.append(local_reg.get())
        data_registro_lista.append(f'{dia}/{mes}/{ano}')

        # criação do histórico
        hist1 = cod_equip_registro_lista[0] + "  -  " + \
                cod_pessoa_registro_lista[0] + "  -  " + \
                local_registro_lista[0] + "  -  " + \
                data_registro_lista[0]
        lb6.configure(text=hist1, background='Gray', foreground='black')

        if len(cod_equip_registro_lista) > 1:
            hist2 = cod_equip_registro_lista[1] + "  -  " + \
                    cod_pessoa_registro_lista[1] + "  -  " + \
                    local_registro_lista[1] + "  -  " + \
                    data_registro_lista[1]
            lb7.configure(text=hist2, background='Gray', foreground='black')
        else:
            hist2 = ''

        if len(cod_equip_registro_lista) > 2:
            hist3 = cod_equip_registro_lista[2] + "  -  " + \
                    cod_pessoa_registro_lista[2] + "  -  " + \
                    local_registro_lista[2] + "  -  " + \
                    data_registro_lista[2]
            lb8.configure(text=hist3, background='Gray', foreground='black')
        else:
            hist3 = ''

        if len(cod_equip_registro_lista) > 3:
            hist4 = cod_equip_registro_lista[3] + "  -  " + \
                    cod_pessoa_registro_lista[3] + "  -  " + \
                    local_registro_lista[3] + "  -  " + \
                    data_registro_lista[3]
            lb9.configure(text=hist4, background='Gray', foreground='black')
        else:
            hist4 = ''

        if len(cod_equip_registro_lista) > 4:
            hist5 = cod_equip_registro_lista[4] + "  -  " + \
                    cod_pessoa_registro_lista[4] + "  -  " + \
                    local_registro_lista[4] + "  -  " + \
                    data_registro_lista[4]
            lb10.configure(text=hist5, background='Gray', foreground='black')
        else:
            hist5 = ''

        cod_equip_reg.delete(0, "end")
        cod_pessoa_reg.delete(0, "end")
        local_reg.set('')
        btn12.configure(command=comando_btn12)
        btn13.configure(command=comando_btn13)

    btn11 = Button(app6, text='Carregar Dados', command=comando_btn61)
    btn11.place(x=165, y=225, height=40, width=170)

    # histórico dos ultimos 5 registros

    lb6 = Label(app6, text='', background='silver', foreground='black')
    lb6.place(x=50, y=280, height=20, width=400)

    lb7 = Label(app6, text='', background='silver', foreground='black')
    lb7.place(x=50, y=310, height=20, width=400)

    lb8 = Label(app6, text='', background='silver', foreground='black')
    lb8.place(x=50, y=340, height=20, width=400)

    lb9 = Label(app6, text='', background='silver', foreground='black')
    lb9.place(x=50, y=370, height=20, width=400)

    lb10 = Label(app6, text='', background='silver', foreground='black')
    lb10.place(x=50, y=400, height=20, width=400)

    def comando_btn12():
        if len(cod_equip_registro_lista) == 0:
            return msg_erro_cadastro_apag()

        atualizar_plan_registro()
        messagebox.showinfo('Cadastrado', f'Os registros foi inserido com sucesso!!')
        app6.destroy()

    def comando_btn13():
        cod_equip_registro_lista.clear()
        cod_pessoa_registro_lista.clear()
        local_registro_lista.clear()
        data_registro_lista.clear()

        lb6.configure(text='Preencha os dados novamente', background='yellow', foreground='black')
        lb7.configure(text='', background='Silver', foreground='black')
        lb8.configure(text='', background='Silver', foreground='black')
        lb9.configure(text='', background='Silver', foreground='black')
        lb10.configure(text='', background='Silver', foreground='black')
        cod_equip_reg.delete(0, "end")
        cod_pessoa_reg.delete(0, "end")
        local_reg.set('')
        btn13.configure(command=msg_erro_apagar)

    btn12 = Button(app6, text='Cadastrar', command=msg_erro_cadastro)
    btn12.place(x=50, y=440, height=40, width=150)

    btn13 = Button(app6, text='Apagar', command=msg_erro_apagar)
    btn13.place(x=300, y=440, height=40, width=150)

    app6.mainloop()


if __name__ == '__main__':
    cadastro_registro()
