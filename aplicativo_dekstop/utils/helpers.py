def carregar_arquivo(caminho):
    with open(caminho, "r", encoding="utf-8") as f:
        return f.read()
