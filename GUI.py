from customtkinter import *
import customtkinter as ctk
from PIL import Image
from matplotlib.figure import Figure
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from botoes import Calculador, Word
from os import remove
        
class GUI:

    def __init__(self):
        
        # Configuração da janela
        janela = CTk()
        janela.title("Calculador de Momentos de Inércia")
        ctk.set_appearance_mode("dark") # pode ser light
        janela.geometry("900x530")
        janela.resizable(width=False, height=False)

        # Frames
        frame1 = ctk.CTkFrame(master = janela, width = 550, height = 420)
        frame1.place(x = 10, y = 50)
        self.frame2 = ctk.CTkFrame(master = janela, width = 320, height = 420)
        self.frame2.place(x = 570, y = 50)
        
        texto_d = ctk.CTkLabel(self.frame2, text = "Digite a maior dimensão vertical da seção (mm)")
        texto_d.pack(side = "top", padx = 10, pady = 10)
        self.caixa_d = ctk.CTkEntry(self.frame2)
        self.caixa_d.pack(side = "top", pady = 2)


        # Função do DropBox
        def escolha(x):
            if x == "Sair":
                janela.destroy()
            if x == "Abrir...":
                foto = filedialog.askopenfilename(filetypes = [("PNG files", "*.png")])
                if foto:
                    self.imagem = Image.open(foto)
                    self.ax.clear()
                    self.ax.imshow(self.imagem)
                    self.canvas.draw()
                    self.botao.configure(state = "normal")

        self.fig = Figure(figsize=(5,3.5), dpi = 100)
        self.ax = self.fig.add_subplot()
        self.canvas = FigureCanvasTkAgg(self.fig, frame1)
        self.canvas.get_tk_widget().pack(padx = 20, pady = 20)

        # Drop Box
        combo_box = ctk.CTkComboBox(janela, values = ["Abrir...", "Sair"], command = escolha)
        combo_box.place(x = 10, y = 10)
        combo_box.set("Arquivo")

        # Botão
        self.botao = CTkButton(janela, text = "Calcular", command = self.processar, state = "disabled")
        self.botao.place(x = 390, y = 485)

        # Iniciar a interface
        janela.mainloop()

    def processar(self):

        D = float(self.caixa_d.get())

        calculador = Calculador()
        img_largura, img_altura, Y_linha, I, posicao_linha = calculador.processa(self.imagem, D)
        
        self.ax.axhline(y=posicao_linha, color = "r", linestyle = "--")
        self.ax.imshow(self.imagem)
        self.canvas.draw()

        texto_img_largura = ctk.CTkLabel(self.frame2, text = f"A largura da imagem é: {img_largura} mm")
        texto_img_largura.pack(pady = 14)
        texto_img_altura =ctk.CTkLabel(self.frame2, text = f"A altura da imagem é:  {img_altura} mm")
        texto_img_altura.pack(pady = 14)
        texto_Y_linha = ctk.CTkLabel(self.frame2, text = f"A altura da linha do CG é: {Y_linha} mm")
        texto_Y_linha.pack(pady = 14)
        texto_I = ctk.CTkLabel(self.frame2, text = f"O momento de inércia é: {I} mm^4")
        texto_I.pack(pady = 13)

        word = Word()
        def gerar_docx():
            self.imagem.save("imagem.png")
            self.fig.savefig("fig.png", dpi = 200)
            word.converte(img_largura, img_altura, Y_linha, I, D)
        botao_docx = CTkButton(self.frame2, text = "Gerar documento DOCX", command = gerar_docx)
        botao_docx.pack(pady = 30)


gui = GUI()