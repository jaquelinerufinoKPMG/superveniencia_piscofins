import pyautogui
import time
 
print("Posicione o mouse na tela. Pressione Ctrl+C para interromper.")
 
try:
    while True:
        # Captura a posição atual do mouse
        x, y = pyautogui.position()
        print(f"Posição do mouse: X={x}, Y={y}", end="\r")
        time.sleep(0.1)  # Atualiza a cada 0.1 segundos
except KeyboardInterrupt:
    print("\nFinalizado.")