import customtkinter as ctk
from controllers.home_controller import HomeController


class HomeView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)


        self.controller = HomeController(self)

        ctk.CTkLabel(self, text="Bem-Vindo").pack(pady=10)

        ctk.CTkButton(self, text="Juntar Detalamento", command=self.controller.ir_para_merge).pack(pady=10)

