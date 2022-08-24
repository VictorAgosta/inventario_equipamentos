from tkinter import ttk
import pandas as pd
from tkinter import *
from tkinter import messagebox
from datetime import date
import tkinter.font as tk_font
from mensagens import msg_erro_cadastro_apag, msg_erro_cadastro, msg_erro_apagar


######################################### Cadastro mobiles #########################################

# criando as variaveis de listas que receberão os dados

cod_mobile_lista = []
nome_mobile_lista = []
marca_mobile_lista = []
modelo_mobile_lista = []
cod_chip_mobile_lista = []


# função para atualizar a tabela mobile


def atualizar_plan_mobile():
    plan0_mobile = pd.read_excel('Inventario_mobile.xlsx', sheet_name='mobile')

    # criação do dicionário
    dicio_mobile = {'cod_mobile': cod_mobile_lista, 'objeto': nome_mobile_lista,
                    'marca': marca_mobile_lista, 'modelo': modelo_mobile_lista, 'cod_chip': cod_chip_mobile_lista}
    df_mobile = pd.DataFrame(dicio_mobile)

    # atualização da planilha
    plan1_mobile = pd.concat([plan0_mobile, df_mobile], ignore_index=True)
    plan1_mobile.to_excel('Inventario_mobile.xlsx', sheet_name='mobile', index=False)


# função da interface para preenchimento dos dados


def cadastro_mobile(cod_mob, obj_mob):
    plan0_chip = pd.read_excel('Inventario_chip.xlsx', sheet_name='chip')

    # criação da interface cadastro pessoas
    app3 = Tk()
    app3.title('Cadastro Mobiles')
    app3.geometry('500x460')
    app3.configure(background='Silver')

    font_titulo = tk_font.Font(family="Lucida Grande", size=15)

    # Titulo tela

    lb0 = Label(app3, text='Cadastro dos Equipamentos Mobile:', background='Silver',
                foreground='black', font=font_titulo)

    lb0.place(x=0, y=10, height=30, width=500)

    lb1 = Label(app3, text='Código Mobile', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=60, height=20, width=100)

    nome_mobile = Label(app3, text=f'{cod_mob}', background='Silver', foreground='black', anchor=W)
    nome_mobile.place(x=120, y=60, height=20, width=350)

    # nome mobile:
    lb2 = Label(app3, text='Equipamento:', background='Silver', foreground='black', anchor=W)
    lb2.place(x=10, y=90, height=20, width=100)

    nome_mobile = Label(app3, text=f'{obj_mob}', background='Silver', foreground='black', anchor=W)
    nome_mobile.place(x=120, y=90, height=20, width=350)

    # marca
    lb3 = Label(app3, text='Marca:', background='Silver', foreground='black', anchor=W)
    lb3.place(x=10, y=120, height=20, width=100)

    marca_mobile = Entry(app3)
    marca_mobile.place(x=120, y=120, height=20, width=350)

    # modelo
    lb4 = Label(app3, text='Modelo:', background='Silver', foreground='black', anchor=W)
    lb4.place(x=10, y=150, height=20, width=100)

    modelo_mobile = Entry(app3)
    modelo_mobile.place(x=120, y=150, height=20, width=350)

    # cod_chip
    lb5 = Label(app3, text='Código do Chip:', background='Silver', foreground='black', anchor=W)
    lb5.place(x=10, y=180, height=20, width=100)

    cod_chip_mobile = Entry(app3)
    cod_chip_mobile.place(x=120, y=180, height=20, width=210)

    # comando para exibir numero do chip atrelado ao codigo
    def comando_btn34():
        codigo_chip = cod_chip_mobile.get()
        detectar_chip_ind = plan0_chip[plan0_chip['cod_chip'] == codigo_chip].index.tolist()
        if not detectar_chip_ind:
            return num_chip_mob.configure(text='Chip não cadastrado')
        numero_chip = plan0_chip.iloc[detectar_chip_ind[0]]['numero']
        num_chip_mob.configure(text=numero_chip)

    # Botão para exibir numero do chip atrelado ao codigo
    btn34 = Button(app3, text='Exibir nº do CHIP', command=comando_btn34)
    btn34.place(x=350, y=180, height=20, width=120)

    # numero chip
    lb6 = Label(app3, text='Número do Chip:', background='Silver', foreground='black', anchor=W)
    lb6.place(x=10, y=210, height=20, width=100)

    num_chip_mob = Label(app3, text=' ', background='Silver', foreground='black', anchor=W)
    num_chip_mob.place(x=120, y=210, height=20, width=350)

    def comando_btn31():
        cod_mobile_lista.append(cod_mob)
        nome_mobile_lista.append(obj_mob)
        marca_mobile_lista.append(marca_mobile.get())
        modelo_mobile_lista.append(modelo_mobile.get())
        cod_chip_mobile_lista.append(cod_chip_mobile.get())

        hist = cod_mobile_lista[0] + " - " + \
               nome_mobile_lista[0] + " - " + \
               marca_mobile_lista[0] + " - " + \
               modelo_mobile_lista[0] + " - " + \
               cod_chip_mobile_lista[0]

        lb7.configure(text=hist, background='Gray')
        btn32.configure(command=comando_btn32)
        btn33.configure(command=comando_btn33)

    btn31 = Button(app3, text='Carregar Dados', command=comando_btn31)
    btn31.place(x=175, y=250, height=40, width=150)

    lb7 = Label(app3, text=' ', background='silver', foreground='black')
    lb7.place(x=50, y=320, height=40, width=400)

    def comando_btn32():
        if len(cod_mobile_lista) == 0:
            return msg_erro_cadastro_apag()

        atualizar_plan_mobile()
        messagebox.showinfo('Cadastrado', f'O mobile {cod_mob} foi inserido com sucesso!!\n'
                                          'Finalize o cadastro do equipamento na tela de Cadastro de Equipamentos.')
        app3.destroy()

    def comando_btn33():
        cod_mobile_lista.clear()
        nome_mobile_lista.clear()
        marca_mobile_lista.clear()
        modelo_mobile_lista.clear()
        cod_chip_mobile_lista.clear()
        lb5.configure(text='Preencha os dados novamente', background='yellow', foreground='black')
        marca_mobile.delete(0, "end")
        modelo_mobile.delete(0, "end")
        cod_chip_mobile.delete(0, "end")
        btn33.configure(command=msg_erro_apagar)

    btn32 = Button(app3, text='Cadastrar', command=msg_erro_cadastro)
    btn32.place(x=50, y=390, height=40, width=150)

    btn33 = Button(app3, text='Apagar', command=msg_erro_apagar)
    btn33.place(x=300, y=390, height=40, width=150)


    app3.mainloop()

if __name__ == '__main__':
    cadastro_mobile('MO10', 'celular')
