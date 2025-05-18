from customtkinter import *
import customtkinter as ctk
from robot import *
from PIL import Image

#adicionando Campos
espaco = " "
multilploespaco = 42*espaco

ctk.set_appearance_mode('dark')
ctk.set_default_color_theme("blue")


#janela = ctk.CTk(fg_color=("white"))
janela = ctk.CTk(fg_color=("#131719"))
janela.title("NexusBOT")

caminho_icone = os.path.abspath("robot.png")
img = Image.open(caminho_icone)
img.save("robo.ico", format="ICO")
janela.iconbitmap("robo.ico")


janela.minsize(300, 300) 
janela.maxsize(300, 300)

janela.grid_columnconfigure(0, weight=1)
janela.grid_columnconfigure(1, weight=1)
janela.grid_columnconfigure(2, weight=1)


janela.grid_rowconfigure(0, weight=1)
janela.grid_rowconfigure(1, weight=1)
janela.grid_rowconfigure(2, weight=1)
janela.grid_rowconfigure(3, weight=1)
janela.grid_rowconfigure(4, weight=1)


texto_Titulo = ctk.CTkLabel(janela, text="NexusBOT", font=("ARIAL", 30, "bold"),text_color="white", fg_color=("#168BDA"), corner_radius=5)
texto_Titulo.grid(row=0, column=1)

image_robo = ctk.CTkImage(light_image=Image.open("robot.png"), size=(94, 94))
label_robo = ctk.CTkLabel(janela, image=image_robo, text="")
label_robo.grid(row=1, column=1)

def clicou_navegador():
    if botaoopc1._clicked:
        CTkMessagebox( title="Sucesso!", message="Operação concluída com sucesso!", icon="check", option_1="OK")
clicou_navegador

botaoopc1 = ctk.CTkButton(janela, text="  ABRIR SIRESP  ", command=iniciar_siresp, font=("Verdana", 11, "bold"), fg_color=("#2E1E76") )
botaoopc1.grid(row=2, column=1)

botaoopc2 = ctk.CTkButton(janela, text="  INICIAR ROBOT  ", command=Iniciar_robot, fg_color=("#2E1E76"), font=("Verdana", 11, "bold"))
botaoopc2.grid(row=3, column=1)

botaoopc3 = ctk.CTkButton(janela, text="IMPORTAR PLANILHA", command=importar_planilha,  fg_color=("#2E1E76"), font=("Verdana", 11, "bold"))
botaoopc3.grid(row=4, column=1)

janela.mainloop()