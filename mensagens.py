from tkinter import messagebox

############################################ Mensagens de Erro ############################################

def msg_erro_cadastro():
    messagebox.showwarning("Erro Cadastro", "Preencha os campos e clique em carregar dados,"
                                            " visualize se os dados estão 'OK' então clique no botão para"
                                            " cadastrar novamente")


def msg_erro_apagar():
    messagebox.showwarning("Erro", "Não há dados para apagar!")


def msg_erro_cadastro_apag():
    messagebox.showwarning("Erro", "Não há dados para cadastrar!")

