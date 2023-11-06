import tkinter as tk
from tkinter import filedialog
import pandas as pd

def fechar_janela():
    janela.quit()

def center_window(win, width=350, height=300):
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)

    win.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

def carregar_arquivo():
    janela.lift()  # Eleva a janela principal para a frente
    filepath = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
    if filepath:
        global df
        df = pd.read_excel(filepath)
        exibir_janela_resultados()

def voltar_e_mostrar_principal(janela_resultados):
    janela_resultados.destroy()
    janela.deiconify()  # Traz a janela principal de volta à frente

def exibir_janela_resultados():
    janela_resultados = tk.Toplevel(janela)
    janela_resultados.title("Resultados")
    janela_resultados.geometry("1280x720")

    global texto_resultados

    frame_botoes = tk.Frame(janela_resultados)
    frame_botoes.pack(side=tk.TOP, pady=10)

    botao_carregar_outro_arquivo = tk.Button(frame_botoes, text="Carregar Outro Arquivo", command=carregar_arquivo, bg="green", fg="white")
    botao_carregar_outro_arquivo.pack(side=tk.LEFT, padx=5)
    botao_voltar = tk.Button(frame_botoes, text="Voltar", command=lambda: voltar_e_mostrar_principal(janela_resultados), bg="blue", fg="white")
    botao_voltar.pack(side=tk.LEFT, padx=5)

    texto_resultados = tk.Text(janela_resultados, width=105, height=50)
    texto_resultados.insert(tk.END, df.to_string(index=True, justify='center'))
    texto_resultados.tag_configure("center", justify='center')
    texto_resultados.tag_add("center", "1.0", "end")
    texto_resultados.pack(expand=True, fill=tk.BOTH)
    texto_resultados.config(wrap="none")
    
    scrollbar_horizontal = tk.Scrollbar(janela_resultados, orient="horizontal", command=texto_resultados.xview)
    scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)
    texto_resultados.configure(xscrollcommand=scrollbar_horizontal.set)
    scrollbar_vertical = tk.Scrollbar(janela_resultados, orient="vertical", command=texto_resultados.yview)
    scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)
    texto_resultados.configure(yscrollcommand=scrollbar_vertical.set)


janela = tk.Tk()
janela.title("Projeto Dashboard")

center_window(janela)

rotulo_titulo = tk.Label(janela, text="LEITOR DE EXCEL")
rotulo_titulo.pack()
rotulo_titulo.place(relx=0.5, rely=0.3, anchor="center")

botao_carregar = tk.Button(janela, text="Selecionar Arquivo", command=carregar_arquivo, bg="green", fg="white")
botao_carregar.pack()
botao_carregar.place(relx=0.5, rely=0.5, anchor="center")

botao_sair = tk.Button(janela, text="SAIR", command=fechar_janela, bg="red", fg="white")
botao_sair.pack()
botao_sair.place(relx=1, rely=1, anchor="se")

janela.mainloop()
