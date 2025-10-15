class UserModel:
    def __init__(self, nome):
        self.nome = nome

    def saudacao(self):
        return f"Olá, {self.nome}!"
