import pyautogui
import time
from PIL import Image

time.sleep(2)

# 1. Clique inicial na posição
x_inicial, y = 78, 545

# 3. Espera meio segundo
time.sleep(0.5)

# 4. Segundo clique
py.click(x_inicial, y)

# 5. Inicia medição de tempo
inicio = time.time()
encontrado = False

# 6. Varredura horizontal
for x in range(79, 143):  # de 79 até 142 inclusive
    pixel_color = py.pixel(x, y)
    r, g, b = pixel_color
    print(f"Pixel em ({x},{y}): {pixel_color}")

    if (r, g, b) == (0, 0, 0):
        print("passe direto")
        encontrado = True
        break
    elif g in [127, 128] and g > r and g > b:
        print("tudo ok")
        encontrado = True
        break
    else:
        print("mais uma vez")

# 7. Finaliza medição de tempo
fim = time.time()
duracao = fim - inicio

if encontrado:
    print(f"Encontrado em {duracao:.2f} segundos.")
else:
    print("Não encontrou nenhuma cor válida entre x = 79 e x = 142.")
    print(f"Tempo total de busca: {duracao:.2f} segundos.")
