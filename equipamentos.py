from tkinter import ttk
import pandas as pd
from tkinter import *
from computadores import cadastro_computador
from mobiles import cadastro_mobile
from chips import cadastro_chip

######################################### Cadastro de Equipamentos #########################################

# criando as variaveis de listas que receberão os dados
cod_equip_lista = []
nome_equip_lista = []
marca_equip_lista = []
tipo_equip_lista = []
status_equip_lista = []


# função para atualizar a tabela equipamentos


def atualizar_plan_equipamento():
    plan0_equip = pd.read_excel('Inventario_equipamento.xlsx', sheet_name='equipamento')

    # criação do dicionário
    dicio_equip = {'cod_equipamento': cod_equip_lista, 'equipamento': nome_equip_lista, 'marca': marca_equip_lista,
                   'tipo': tipo_equip_lista, 'status': status_equip_lista}
    df_equip = pd.DataFrame(dicio_equip)

    # atualização da planilha
    plan1_pessoas = pd.concat([plan0_equip, df_equip], ignore_index=True)
    plan1_pessoas.to_excel('Inventario_equipamento.xlsx', sheet_name='equipamento', index=False)

    cod_equip_lista.clear()
    nome_equip_lista.clear()
    marca_equip_lista.clear()
    tipo_equip_lista.clear()
    status_equip_lista.clear()


def cadastro_equipamento():
    # criação da interface cadastro equipamentos
    app2 = Tk()
    app2.title('Cadastro Equipamentos')
    app2.geometry('500x500')
    app2.configure(background='Silver')

    plan0_equip = pd.read_excel('Inventario_equipamento.xlsx', sheet_name='equipamento')
    plan0_mobile = pd.read_excel('Inventario_mobile.xlsx', sheet_name='mobile')
    plan0_chip = pd.read_excel('Inventario_chip.xlsx', sheet_name='chip')

    # nome:
    lb1 = Label(app2, text='Equipamento:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=30, height=20, width=100)

    nome_equip = ttk.Combobox(app2, values=['All in One', 'Apoio de Mãos (teclado)', 'Apoio de Pés',
                                            'Suporte Gaveteiro', 'Suporte Notebook', 'Mouse C/ Fio',
                                            'Mouse S/ Fio', 'Mousepad C/ Apoio', 'Mousepad S/ Apoio',
                                            'Teclado C/ Fio', 'Teclado S/ Fio', 'Desktop', 'Notebook', 'Monitor',
                                            'Outros*', 'Headset', 'Fonte Celular', 'Cabo USB Celular',
                                            'Celular', 'Chip', 'Tablet'], state="readonly")

    def cadastro_carac(event):
        nome = nome_equip.get()
        if nome == 'All in One':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'All in One'].count()
            contagem = contagem_cods[0]
            cod = f'AIOINM{contagem + 1}'
            cadastro_computador(cod_comp=cod)

        elif nome == 'Desktop':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Desktop'].count()
            contagem = contagem_cods[0]
            cod = f'DSKINM{contagem + 1}'
            cadastro_computador(cod_comp=cod)

        elif nome == 'Notebook':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Notebook'].count()
            contagem = contagem_cods[0]
            cod = f'NBKINM{contagem + 1}'
            cadastro_computador(cod_comp=cod)

        elif nome == 'Celular':
            contagem_cods = plan0_mobile.loc[plan0_mobile['objeto'] == 'Celular'].count()
            contagem = contagem_cods[0]
            cod = f'MB{contagem + 1}'
            obj = nome
            cadastro_mobile(cod_mob=cod, obj_mob=obj)

        elif nome == 'Tablet':
            contagem_cods = plan0_mobile.loc[plan0_mobile['objeto'] == 'Tablet'].count()
            contagem = contagem_cods[0]
            cod = f'MB{contagem + 1}'
            obj = nome
            cadastro_mobile(cod_mob=cod, obj_mob=obj)

        elif nome == 'Chip':
            contagem_cods = plan0_chip.count()
            contagem = contagem_cods[0]
            cod = f'CH{contagem + 1}'
            cadastro_chip(cod_chip=cod)

    nome_equip.place(x=120, y=30, height=20, width=350)

    nome_equip.bind("<<ComboboxSelected>>", cadastro_carac)

    # marca
    lb2 = Label(app2, text='Marca:', background='Silver', foreground='black', anchor=W)
    lb2.place(x=10, y=60, height=20, width=100)

    marca_equip = Entry(app2)
    marca_equip.place(x=120, y=60, height=20, width=350)

    # tipo
    lb3 = Label(app2, text='Equipamento (outros)*', background='Silver', foreground='black', anchor=W)
    lb3.place(x=10, y=90, height=20, width=100)

    tipo_equip = Entry(app2)
    tipo_equip.place(x=120, y=90, height=20, width=350)

    def comando_btn21():
        equip = nome_equip.get()

        if equip == 'All in One':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'All in One'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'AIOINM{contagem + 1}')

        elif equip == 'Apoio de Mãos (teclado)':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Apoio de Mãos (teclado)'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'APM{contagem + 1}')

        elif equip == 'Apoio de Pés':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Apoio de Pés'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'APP{contagem + 1}')

        elif equip == 'Suporte Gaveteiro':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Suporte Gaveteiro'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'SPG{contagem + 1}')

        elif equip == 'Suporte Notebook':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Suporte Notebook'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'SPN{contagem + 1}')

        elif equip == 'Mouse C/ Fio':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Mouse C/ Fio'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'MSC{contagem + 1}')

        elif equip == 'Mouse S/ Fio':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Mouse S/ Fio'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'MSS{contagem + 1}')

        elif equip == 'Mousepad C/ Apoio':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Mousepad C/ Apoio'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'MPC{contagem + 1}')

        elif equip == 'Mousepad S/ Apoio':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Mousepad S/ Apoio'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'MPS{contagem + 1}')

        elif equip == 'Teclado C/ Fio':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Teclado C/ Fio'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'TCC{contagem + 1}')

        elif equip == 'Teclado S/ Fio':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Teclado S/ Fio'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'TCS{contagem + 1}')

        elif equip == 'Desktop':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Desktop'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'DSKINM{contagem + 1}')

        elif equip == 'Notebook':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Notebook'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'NBKINM{contagem + 1}')

        elif equip == 'Monitor':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Monitor'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'MO{contagem + 1}')

        elif equip == 'Headset':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Headset'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'HS{contagem + 1}')

        elif equip == 'Fonte Celular':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Fonte Celular'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'FTC{contagem + 1}')

        elif equip == 'Cabo USB Celular':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Cabo USB Celular'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'CUC{contagem + 1}')

        elif equip == 'Outros*':
            contagem_cods = plan0_equip.loc[plan0_equip['equipamento'] == 'Equipamento TI'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'ETI{contagem + 1}')

        elif equip == 'Celular':
            contagem_cods = plan0_mobile.loc[plan0_mobile['objeto'] == 'Celular'].count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'MB{contagem + 1}')

        elif equip == 'Chip':
            contagem_cods = plan0_chip.count()
            contagem = contagem_cods[0]
            cod_equip_lista.append(f'CH{contagem + 1}')

        nome_equip_lista.append(nome_equip.get())
        marca_equip_lista.append(marca_equip.get())
        tipo_equip_lista.append(tipo_equip.get())
        status_equip_lista.append('Estoque')

        hist1 = cod_equip_lista[0] + " " + \
                nome_equip_lista[0] + " " + \
                marca_equip_lista[0] + " " + \
                tipo_equip_lista[0] + " " + \
                status_equip_lista[0]

        if len(cod_equip_lista) > 1:
            hist2 = cod_equip_lista[1] + " " + \
                    nome_equip_lista[1] + " " + \
                    marca_equip_lista[1] + " " + \
                    tipo_equip_lista[1] + " " + \
                    status_equip_lista[1]
        else:
            hist2 = ''

        if len(cod_equip_lista) > 2:
            hist3 = cod_equip_lista[2] + " " + \
                    nome_equip_lista[2] + " " + \
                    marca_equip_lista[2] + " " + \
                    tipo_equip_lista[2] + " " + \
                    status_equip_lista[2]
        else:
            hist3 = ''

        if len(cod_equip_lista) > 3:
            hist4 = cod_equip_lista[3] + " " + \
                    nome_equip_lista[3] + " " + \
                    marca_equip_lista[3] + " " + \
                    tipo_equip_lista[3] + " " + \
                    status_equip_lista[3]
        else:
            hist4 = ''

        if len(cod_equip_lista) > 4:
            hist5 = cod_equip_lista[4] + " " + \
                    nome_equip_lista[4] + " " + \
                    marca_equip_lista[4] + " " + \
                    tipo_equip_lista[4] + " " + \
                    status_equip_lista[4]
        else:
            hist5 = ''

        lb5.configure(text=hist1)
        lb6.configure(text=hist2)
        lb7.configure(text=hist3)
        lb8.configure(text=hist4)
        lb9.configure(text=hist5)

        nome_equip.set('')
        marca_equip.delete(0, "end")
        tipo_equip.delete(0, "end")

    btn11 = Button(app2, text='Carregar Dados', command=comando_btn21)
    btn11.place(x=100, y=170, height=40, width=150)

    # botão para atualizar a planilha com os dados carregados
    btn12 = Button(app2, text='Atualizar Planilha', command=atualizar_plan_equipamento)
    btn12.place(x=300, y=170, height=40, width=150)

    lb5 = Label(app2, text='', background='Silver', foreground='black', anchor=W)
    lb5.place(x=50, y=220, height=20, width=300)

    lb6 = Label(app2, text='', background='Silver', foreground='black', anchor=W)
    lb6.place(x=50, y=250, height=20, width=300)

    lb7 = Label(app2, text='', background='Silver', foreground='black', anchor=W)
    lb7.place(x=50, y=280, height=20, width=300)

    lb8 = Label(app2, text='', background='Silver', foreground='black', anchor=W)
    lb8.place(x=50, y=310, height=20, width=300)

    lb9 = Label(app2, text='', background='Silver', foreground='black', anchor=W)
    lb9.place(x=50, y=340, height=20, width=300)

    app2.mainloop()


if __name__ == '__main__':
    cadastro_equipamento()
