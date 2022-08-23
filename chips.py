
######################################### Cadastro Chips #########################################

# criando as variaveis de listas que receberão os dados

cod_chip_lista = []
numero_chip_lista = []


# função para atualizar a tabela mobile


def atualizar_plan_chip():
    plan0_chip = pd.read_excel('Inventario_chip.xlsx', sheet_name='chip')

    # criação do dicionário
    dicio_chip = {'cod_chip': cod_chip_lista, 'numero': numero_chip_lista}
    df_chip = pd.DataFrame(dicio_chip)

    # atualização da planilha
    plan1_chip = pd.concat([plan0_chip, df_chip], ignore_index=True)
    plan1_chip.to_excel('Inventario_chip.xlsx', sheet_name='chip', index=False)

    # limpar as listas para caso queira atualizar com mais valores a planilha
    cod_chip_lista.clear()
    numero_chip_lista.clear()


# função da interface para preenchimento dos dados


def cadastro_chip(cod_chip):
    # criação da interface cadastro pessoas
    app4 = Toplevel()
    app4.title('Cadastro Chip')
    app4.geometry('500x500')
    app4.configure(background='Silver')

    # Codigo chip:
    lb0 = Label(app4, text='Código do chip:', background='Silver', foreground='black', anchor=W)
    lb0.place(x=10, y=30, height=20, width=100)

    numero_chip = Label(app4, text=f'{cod_chip}', background='Silver', foreground='black', anchor=W)
    numero_chip.place(x=120, y=30, height=20, width=350)

    # numero chip:
    lb1 = Label(app4, text='Número do chip:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=60, height=20, width=100)

    numero_chip = Entry(app4)
    numero_chip.place(x=120, y=60, height=20, width=350)

    def comando_btn41():
        cod_chip_lista.append(cod_chip)
        numero_chip_lista.append(numero_chip.get())
        numero_chip.delete(0, "end")

        hist = cod_chip_lista[0] + " - " + \
               numero_chip_lista[0]

        lb2.configure(text=hist, background='Black', foreground='white')
        btn42.configure(command=comando_btn42)
        btn43.configure(command=comando_btn43)

    btn41 = Button(app4, text='Cadastrar Dados', command=comando_btn41)
    btn41.place(x=100, y=90, height=40, width=150)

    lb2 = Label(app4, text=' ', background='Silver', foreground='black', anchor=W)
    lb2.place(x=50, y=250, height=40, width=400)

    def comando_btn42():
        if len(cod_chip_lista) == 0:
            return msg_erro_cadastro_apag()

        atualizar_plan_chip()
        messagebox.showinfo('Cadastrado', f'O chip {cod_chip} foi inserido com sucesso!!\n'
                                          'Finalize o cadastro do equipamento na tela de Cadastro de Equipamentos.')
        app4.destroy()

    def comando_btn43():
        cod_chip_lista.clear()
        numero_chip_lista.clear()

        lb2.configure(text='Preencha os dados novamente', background='yellow', foreground='black')
        numero_chip.delete(0, "end")

    btn42 = Button(app4, text='Cadastrar', command=msg_erro_cadastro)
    btn42.place(x=50, y=350, height=40, width=150)

    btn43 = Button(app4, text='Apagar', command=msg_erro_apagar)
    btn43.place(x=275, y=350, height=40, width=150)
