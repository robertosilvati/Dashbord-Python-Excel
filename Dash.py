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
    janela.lift()
    filepath = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
    if filepath:
        global df
        df = pd.read_excel(filepath)
        exibir_janela_resultados()

def voltar_e_mostrar_principal(janela_resultados):
    janela_resultados.destroy()
    janela.deiconify()

def atualizar_informacoes():
    texto_resultados.delete("1.0", tk.END)
    texto_resultados.insert(tk.END, df.to_string(index=True, justify='center'))

def adaptar_geometria(janela):
    largura = janela.winfo_screenwidth()
    altura = janela.winfo_screenheight()

    proporcao_largura = 0.8
    proporcao_altura = 0.8

    largura_janela = int(largura * proporcao_largura)
    altura_janela = int(altura * proporcao_altura)

    posicao_x = (largura - largura_janela) // 2
    posicao_y = (altura - altura_janela) // 2

    janela.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")
    janela.minsize(width=667, height=375)

def exibir_janela_resultados():
    janela_resultados = tk.Toplevel(janela)
    janela_resultados.title("Resultados")
    adaptar_geometria(janela_resultados)

    global texto_resultados

    frame_topo = tk.Frame(janela_resultados)
    frame_topo.pack(side=tk.TOP, pady=10)

    botao_voltar = tk.Button(frame_topo, text="Voltar", command=lambda: voltar_e_mostrar_principal(janela_resultados), bg="green", fg="white")
    botao_voltar.pack(side=tk.LEFT, padx=5)

    frame_baixo = tk.Frame(janela_resultados)
    frame_baixo.pack(side=tk.BOTTOM, pady=10)

    botao_carregar_outro_arquivo = tk.Button(frame_baixo, text="Carregar Outro Arquivo", command=carregar_arquivo, bg="green", fg="white")
    botao_carregar_outro_arquivo.pack(side=tk.BOTTOM, padx=10)

    texto_resultados = tk.Text(janela_resultados, width=105, height=50)
    texto_resultados.insert(tk.END, df.to_string(index=True, justify='center'))
    texto_resultados.tag_configure("left", justify='left')
    scrollbar_horizontal = tk.Scrollbar(janela_resultados, orient="horizontal", command=texto_resultados.xview)
    scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)
    texto_resultados.configure(xscrollcommand=scrollbar_horizontal.set)
    texto_resultados.insert(tk.END, df.to_string(index=True, justify='center'))
    texto_resultados.pack(expand=True, fill=tk.BOTH)
    texto_resultados.config(wrap="none")

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
