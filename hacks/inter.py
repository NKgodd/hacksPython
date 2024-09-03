import tkinter as tk
from tkinter import messagebox
import time

def show_ad():
    # Cria a janela principal
    root = tk.Tk()
    root.title("Aviso do Sistema")
    
    # Define o tamanho da janela
    root.geometry("400x600")
    
    # Adiciona uma etiqueta com o texto do anúncio
    label = tk.Label(root, text="Sistema Comprometido", padx=20, pady=20)
    label.pack()
    
    # Adiciona um botão para fechar o anúncio
    button = tk.Button(root, text="Fechar", command=root.destroy)
    button.pack()
    
    # Exibe a janela
    root.mainloop()

def main():
    while True:
        show_ad()
        time.sleep(5)  # Espera 10 segundos 

if __name__ == "__main__":
    main()
