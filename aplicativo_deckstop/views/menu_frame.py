import customtkinter as ctk
from views.detalhamento_frame import Detalhamento
from views.concaternar_frame import Concatenar
from views.excluir_linhas import ExcluirLinhas


class MenuFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.label = ctk.CTkLabel(self, text="Menu Principal")
        self.label.pack(pady=(10,20))

        self.btn_unir_dados = ctk.CTkButton(self, text="Planilhas de Detalhamento", command=lambda: master.show_frame(Detalhamento), width=200)

        self.btn_unir_dados.pack(pady=(0, 5))


        self.btn_editar_dados = ctk.CTkButton(self, text="Editar Dados", width=200)

        self.btn_editar_dados.pack(pady=(0,5))

        self.btn_concatenar_planilhas = ctk.CTkButton(self, text="Concatenar Planilhas", command=lambda: master.show_frame(Concatenar), width=200)

        self.btn_concatenar_planilhas.pack(pady=(0,5))

        self.btn_excluir_linhas = ctk.CTkButton(self, text="Excluir Linhas", width=200, command=lambda: master.show_frame(ExcluirLinhas))

        self.btn_excluir_linhas.pack(pady=(0,5))
        