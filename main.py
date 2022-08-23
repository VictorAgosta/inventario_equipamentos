from tkinter import ttk
import pandas as pd
from tkinter import *
from tkinter import messagebox
from registros import cadastro_registro


############################################ Mensagens de Erro ############################################

def msg_erro_cadastro():
    messagebox.showwarning("Erro Cadastro", "Preencha os campos e clique em carregar dados,"
                                            " visualize se os dados estão 'OK' então clique no botão para"
                                            " cadastrar novamente")


def msg_erro_apagar():
    messagebox.showwarning("Erro", "Não há dados para apagar!")


def msg_erro_cadastro_apag():
    messagebox.showwarning("Erro", "Não há dados para cadastrar!")

######################################### Inativação de Usuários #########################################

cod_pessoa_inativ_lista = []
cod_equips_reg_lista = []


def atualizar_plan_pessoas_inativa():
    plan0_pessoas_inat = pd.read_excel('Inventario_usuario.xlsx', sheet_name='usuario')
    plan0_equip = pd.read_excel('Inventario_equipamento.xlsx', sheet_name='equipamento')
    plan0_reg = pd.read_excel('Inventario_registros.xlsx', sheet_name='registros')

    for usuario in cod_pessoa_inativ_lista:
        detectar_pessoa_ind = plan0_pessoas_inat[plan0_pessoas_inat['cod_usuario'] == usuario].index
        plan0_pessoas_inat['status'].values[detectar_pessoa_ind] = 'Inativo'
        plan_reg_usuario = plan0_reg.loc[plan0_reg['cod_usuario'] == usuario]
        print(plan_reg_usuario)
        for equipamento in plan_reg_usuario['cod_equipamento']:
            cod_equips_reg_lista.append(equipamento)

    for codigo in cod_equips_reg_lista:
        detectar_equip_ind = plan0_equip[plan0_equip['cod_equipamento'] == codigo].index
        plan0_equip['status'].values[detectar_equip_ind] = 'Estoque'

    plan0_equip.to_excel('Inventario_equipamento.xlsx', sheet_name='equipamento', index=False)
    plan0_pessoas_inat.to_excel('Inventario_usuario.xlsx', sheet_name='usuario', index=False)


def inativa_pessoa():
    # criação da interface cadastro equipamentos
    app7 = Toplevel()
    app7.title('Inativação de Usuário')
    app7.geometry('500x200')
    app7.configure(background='Silver')

    # codigo:
    lb1 = Label(app7, text='Código Usuário:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=30, height=20, width=100)

    cod_inat_usuario = Entry(app7)
    cod_inat_usuario.place(x=120, y=30, height=20, width=350)

    def comando_btn71():
        cod_pessoa_inativ_lista.append(cod_inat_usuario.get())
        cod_inat_usuario.delete(0, "end")

    btn11 = Button(app7, text='Carregar Dados', command=comando_btn71)
    btn11.place(x=100, y=80, height=40, width=150)

    # botão para atualizar a planilha com os dados carregados
    btn12 = Button(app7, text='Atualizar Planilha', command=atualizar_plan_pessoas_inativa)
    btn12.place(x=300, y=80, height=40, width=150)

    app7.mainloop()


######################################### Inativação de Equipamento #########################################

cod_equip_inativ_lista = []


def atualizar_plan_equip_inativa():
    plan0_equip_inat = pd.read_excel('Inventario_equipamento.xlsx', sheet_name='equipamento')
    for equipamento in cod_equip_inativ_lista:
        detectar_equip_ind = plan0_equip_inat[plan0_equip_inat['cod_equipamento'] == equipamento].index
        plan0_equip_inat['status'].values[detectar_equip_ind] = 'Descartado'

    plan0_equip_inat.to_excel('Inventario_equipamento.xlsx', sheet_name='equipamento', index=False)


def inativa_equip():
    # criação da interface cadastro equipamentos
    app7 = Toplevel()
    app7.title('Descarte de Equipamento')
    app7.geometry('500x200')
    app7.configure(background='Silver')

    # codigo:
    lb1 = Label(app7, text='Código Equipamento:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=30, height=20, width=150)

    cod_inat_equip = Entry(app7)
    cod_inat_equip.place(x=150, y=30, height=20, width=300)

    def comando_btn71():
        cod_equip_inativ_lista.append(cod_inat_equip.get())
        cod_inat_equip.delete(0, "end")

    btn11 = Button(app7, text='Carregar Dados', command=comando_btn71)
    btn11.place(x=100, y=80, height=40, width=150)

    # botão para atualizar a planilha com os dados carregados
    btn12 = Button(app7, text='Atualizar Planilha', command=atualizar_plan_equip_inativa)
    btn12.place(x=300, y=80, height=40, width=150)

    app7.mainloop()


######################################### Interface Cadastro #########################################

app0 = Tk()
app0.title('Cadastro Inventário')
app0.geometry('500x500')
app0.configure(background='Silver')

# botão pessoas

btn01 = Button(app0, text='Cadastro de Pessoas', command=cadastro_pessoa)
btn01.place(x=150, y=120, height=30, width=200)

# botão equipamentos

btn02 = Button(app0, text='Cadastro dos Equipamentos', command=cadastro_equipamento)
btn02.place(x=150, y=170, height=30, width=200)

# botão registros

btn06 = Button(app0, text='Registros', command=cadastro_registro)
btn06.place(x=100, y=50, height=30, width=300)

# botão inativa pessoas

btn07 = Button(app0, text='Inativar Pessoas', command=inativa_pessoa)
btn07.place(x=100, y=400, height=30, width=300)

# botão descarte equipamento

btn07 = Button(app0, text='Descarte de Equipamento', command=inativa_equip)
btn07.place(x=100, y=450, height=30, width=300)

app0.mainloop()
