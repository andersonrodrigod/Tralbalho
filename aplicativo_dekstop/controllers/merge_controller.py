from tkinter import filedialog, messagebox
from models.merge_model import MergeModel

class MergeController:
    def __init__(self, view):
        self.view = view
        self.file1 = None
        self.file2 = None

    def selecionar_arquivo1(self):
        caminho = filedialog.askopenfilename(title="Selecione o primeiro Excel", filetypes=[("Excel files", "*.xlsx")])
        if caminho:
            self.file1 = caminho
            nome = caminho.split("/")[-1]
            self.view.label_arquivo1.configure(text=f"1º: {nome}", text_color="white")

    def selecionar_arquivo2(self):
        caminho = filedialog.askopenfilename(title="Selecione o segundo Excel", filetypes=[("Excel files", "*.xlsx")])
        if caminho:
            # Evita duplicação: se for o mesmo arquivo, alerta o usuário
            if caminho == self.file1:
                return messagebox.showwarning("Aviso", "Você selecionou o mesmo arquivo duas vezes!")
            self.file2 = caminho
            nome = caminho.split("/")[-1]
            self.view.label_arquivo2.configure(text=f"2º: {nome}", text_color="white")

    def juntar_arquivos(self):
        if not self.file1 or not self.file2:
            return messagebox.showwarning("Aviso", "Selecione os dois arquivos antes de juntar!")

        model = MergeModel(self.file1, self.file2)
        merged_path = model.juntar()
        messagebox.showinfo("Sucesso", f"Arquivos unidos com sucesso!\nSalvo em:\n{merged_path}")
