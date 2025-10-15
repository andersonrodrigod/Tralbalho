import customtkinter as ctk
from controllers.merge_controller import MergeController

class MergeView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.controller = MergeController(self)

        # Botão de voltar
        botao_voltar = ctk.CTkButton(
            self,
            text="←",
            width=30,
            height=30,
            corner_radius=20,
            fg_color="transparent",
            hover_color="#2b2b2b",
            command=lambda: master.show_view("home")
        )
        botao_voltar.place(x=10, y=10)

        # Título
        ctk.CTkLabel(self, text="Juntar Arquivos Excel", font=("Arial", 18, "bold")).pack(pady=20)

        # Botões de seleção
        self.botao_arquivo1 = ctk.CTkButton(self, text="Selecionar 1º Arquivo", command=self.controller.selecionar_arquivo1)
        self.botao_arquivo1.pack(pady=5)

        self.label_arquivo1 = ctk.CTkLabel(self, text="Nenhum arquivo selecionado", text_color="gray")
        self.label_arquivo1.pack()

        self.botao_arquivo2 = ctk.CTkButton(self, text="Selecionar 2º Arquivo", command=self.controller.selecionar_arquivo2)
        self.botao_arquivo2.pack(pady=5)

        self.label_arquivo2 = ctk.CTkLabel(self, text="Nenhum arquivo selecionado", text_color="gray")
        self.label_arquivo2.pack()

        # Botão final para juntar
        ctk.CTkButton(
            self,
            text="Executar",
            width=100,
            command=self.controller.juntar_arquivos
        ).pack(pady=20)
