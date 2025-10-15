import customtkinter as ctk
from views.home_view import HomeView
from views.merge_view import MergeView

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Meu App Estruturado")
        self.geometry("400x300")

        self.current_view = None
        self.show_view("home")

    def show_view(self, view_name):

        if self.current_view:
            self.current_view.destroy()

        if view_name == "home":
            self.current_view = HomeView(self)
        elif view_name == "merge":
            self.current_view = MergeView(self)

        self.current_view.pack(expand=True, fill="both") #type: ignore

