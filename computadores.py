import pandas as pd
from tkinter import *
from tkinter import messagebox
import tkinter.font as tk_font
from mensagens import msg_erro_cadastro_apag, msg_erro_cadastro, msg_erro_apagar

######################################### Cadastro de Computadores #########################################

############################################################################################################
# criando as variaveis de listas que receberão os dados

cod_comp_lista = []
armazenamento_comp_lista = []
ram_comp_lista = []
processador_comp_lista = []

############################################################################################################
# função para atualizar a tabela equipamentos

def atualizar_plan_computadores():
    plan0_comp = pd.read_excel('Inventario_computador.xlsx', sheet_name='computador')

    # criação do dicionário
    dicio_comp = {'cod_computador': cod_comp_lista, 'armazenamento': armazenamento_comp_lista,
                  'ram': ram_comp_lista, 'processador': processador_comp_lista}
    df_comp = pd.DataFrame(dicio_comp)

    # atualização da planilha
    plan1_comp = pd.concat([plan0_comp, df_comp], ignore_index=True)
    plan1_comp.to_excel('Inventario_computador.xlsx', sheet_name='computador', index=False)

    cod_comp_lista.clear()
    armazenamento_comp_lista.clear()
    ram_comp_lista.clear()
    processador_comp_lista.clear()

############################################################################################################
# função para interface da tabela equipamentos

def cadastro_computador(cod_comp):
    # criação da interface cadastro equipamentos
    app5 = Tk()
    app5.title('Cadastro Computadores')
    app5.geometry('500x420')
    app5.configure(background='Silver')

    font_titulo = tk_font.Font(family="Lucida Grande", size=15)

    # Titulo tela

    lb0 = Label(app5, text='Cadastro das Especificações de Computadores:', background='Silver',
                foreground='black', font=font_titulo)

    lb0.place(x=0, y=10, height=30, width=500)

    # codigo:
    lb1 = Label(app5, text='Código:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=60, height=20, width=100)

    codi_comp = Label(app5, text=f'{cod_comp}', background='Silver', foreground='black', anchor=W)
    codi_comp.place(x=120, y=60, height=20, width=350)

    # armazenamento:
    lb1 = Label(app5, text='Armazenamento:', background='Silver', foreground='black', anchor=W)
    lb1.place(x=10, y=90, height=20, width=100)

    armazenamento_comp = Entry(app5)
    armazenamento_comp.place(x=120, y=90, height=20, width=350)

    # ram
    lb2 = Label(app5, text='RAM:', background='Silver', foreground='black', anchor=W)
    lb2.place(x=10, y=120, height=20, width=100)

    ram_comp = Entry(app5)
    ram_comp.place(x=120, y=120, height=20, width=350)

    # processador
    lb3 = Label(app5, text='Processador:', background='Silver', foreground='black', anchor=W)
    lb3.place(x=10, y=150, height=20, width=100)

    processador_comp = Entry(app5)
    processador_comp.place(x=120, y=150, height=20, width=350)

    def comando_btn51():
        cod_comp_lista.append(cod_comp)
        armazenamento_comp_lista.append(armazenamento_comp.get())
        ram_comp_lista.append(ram_comp.get())
        processador_comp_lista.append(processador_comp.get())

        hist = cod_comp_lista[0] + " - " + \
               armazenamento_comp_lista[0] + " - " + \
               ram_comp_lista[0] + " - " + \
               processador_comp_lista[0]

        lb4.configure(text=hist, background='Gray')
        btn52.configure(command=comando_btn52)
        btn53.configure(command=comando_btn53)

    btn11 = Button(app5, text='Carregar Dados', command=comando_btn51)
    btn11.place(x=175, y=210, height=40, width=150)

    lb4 = Label(app5, text=' ', background='Silver', foreground='black')
    lb4.place(x=50, y=280, height=40, width=400)

    def comando_btn52():
        if len(cod_comp_lista) == 0:
            return msg_erro_cadastro_apag()

        atualizar_plan_computadores()
        messagebox.showinfo('Cadastrado', f'O chip {cod_comp} foi inserido com sucesso!!\n'
                                          'Finalize o cadastro do equipamento na tela de Cadastro de Equipamentos.')
        app5.destroy()

    def comando_btn53():
        cod_comp_lista.clear()
        armazenamento_comp_lista.clear()
        ram_comp_lista.clear()
        processador_comp_lista.clear()

        lb4.configure(text='Preencha os dados novamente', background='yellow')
        armazenamento_comp.delete(0, "end")
        ram_comp.delete(0, "end")
        processador_comp.delete(0, "end")
        btn53.configure(command=msg_erro_apagar)

    btn52 = Button(app5, text='Cadastrar Computador', command=msg_erro_cadastro)
    btn52.place(x=50, y=350, height=40, width=150)

    btn53 = Button(app5, text='Apagar Dados', command=msg_erro_apagar)
    btn53.place(x=300, y=350, height=40, width=150)

    app5.mainloop()

if __name__ == '__main__':
    cadastro_computador(cod_comp='CH10')
