class HomeController:
    def __init__(self, view):
        self.view = view

    def botao_clicado(self):
        self.view.label.configure(text="Botão clicado!")

    def ir_para_merge(self):
        self.view.master.show_view("merge")









