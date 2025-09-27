import pyperclip
import pyautogui
from time import sleep
from tkinter import *
import tkinter as tk
import openpyxl
import keyboard
from tkinter import font



#Configurando a Janela
menu = Tk()
menu.title("Cadastro Automatico")
menu.resizable(False,False)
menu.configure(bg='#2C3E50')
menu.attributes('-topmost', True)
menu.overrideredirect(True)

#Dimensão Janela
largura = 300
altura = 200

#Resolução Form

width_screen = menu.winfo_screenwidth()
height_screen = menu.winfo_screenheight()

#Posição Janela
posx = width_screen/2 - largura/2
posy = height_screen/2 - altura/2

#geometry
menu.geometry("%dx%d+%d+%d" %(largura, altura, posx, posy))


# Função para iniciar o arrastar da janela
def iniciar_arrastar(event):
    menu.x_offset = event.x
    menu.y_offset = event.y

# Função para arrastar a janela
def arrastar(event):
    x = menu.winfo_pointerx() - menu.x_offset
    y = menu.winfo_pointery() - menu.y_offset
    menu.geometry(f'+{x}+{y}')
menu.bind("<Button-1>", iniciar_arrastar) 
menu.bind("<B1-Motion>", arrastar) 

def btn_Click(iniciar):
    global executando
    executar()

def executar():
    #workbook é uma variavel que está conectando a planinha do excel com o python
    workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')

    #Ele vai armazenar na variavel a pagina dos produtos puxando o nome da planinha
    pagina_produtos = workbook['Produtos']

    #for feito para ler as linhas da pagina produtos com inicio na linha 2
    for linha in pagina_produtos.iter_rows(min_row=2):
        #armazena o nome do produto x na variavel
        nm_produto = linha[0].value

        #copia o nome armazenado na variavel
        pyperclip.copy(nm_produto)

        #Simula um mouse e vai até as coordenadas desejadas e depois clica
        pyautogui.click(66,245,duration=1)

        #Simula o teclado e nesse caso cola o que está copiado
        pyautogui.hotkey('ctrl','v')

        #Descrição produto
        desc = linha[1].value
        pyperclip.copy(desc)
        pyautogui.click(81,328,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Categoria
        categoria = linha[2].value
        pyperclip.copy(categoria)
        pyautogui.click(75,467,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Codigo produto
        cd_produto = linha[3].value
        pyperclip.copy(cd_produto)
        pyautogui.click(75,551,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Peso
        peso = linha[4].value
        pyperclip.copy(peso)
        pyautogui.click(75,638,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Dimensões
        dimensoes = linha[5].value
        pyperclip.copy(dimensoes)
        pyautogui.click(75,725,duration=1)
        pyautogui.hotkey('ctrl','v')
        
        pyautogui.click(78,774,duration=1)
        #Temporizador
        sleep(3)

        #Preço
        preco = linha[6].value
        pyperclip.copy(preco)
        pyautogui.click(67,263,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Quantidade do estoque
        qnt_estoque = linha[7].value
        pyperclip.copy(qnt_estoque)
        pyautogui.click(65,354,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Validade
        validade = linha[8].value
        pyperclip.copy(validade)
        pyautogui.click(65,433,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Cor
        cor = linha[9].value
        pyperclip.copy(validade)
        pyautogui.click(85,522,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Tamanho
        pyautogui.click(80,614,duration=1)
        tamanho = linha[10].value
        if tamanho == 'Pequeno':
            pyautogui.click(91,640,duration=1)
            
        elif tamanho == 'Médio':
            pyautogui.click(91,674,duration=1)
        
        else:
            pyautogui.click(91,702,duration=1)
            
        #Material
        material = linha[11].value
        pyperclip.copy(material)
        pyautogui.click(77,699,duration=1)
        pyautogui.hotkey('ctrl','v')

        pyautogui.click(87,756,duration=1)
        sleep(3)

        #Fabricante
        fabricante = linha[12].value
        pyperclip.copy(fabricante)
        pyautogui.click(70,305,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Origem
        origem = linha[13].value
        pyperclip.copy(origem)
        pyautogui.click(67,385,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Observação do produto
        observacao = linha[14].value
        pyperclip.copy(observacao)
        pyautogui.click(73,477,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Codigo de barra
        cd_barra = linha[15].value
        pyperclip.copy(cd_barra)
        pyautogui.click(    66,607,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Localização armazem
        localizacao_armazem = linha[16].value
        pyperclip.copy(localizacao_armazem)
        pyautogui.click(71,692,duration=1)
        pyautogui.hotkey('ctrl','v')

        #Confirmar
        pyautogui.click(86,753,duration=1)
        sleep(1)

        #Confirmar inclusão
        pyautogui.click(568,183,duration=1)

        #Repetir cadastro
        pyautogui.click(359,511,duration=1)


def btn_Parar(parar):
    global executando
    executando = False  

def btn_fechar():
    menu.quit()

label1 = Label(menu, text="Bem-Vindo", fg="White", font="Arial 20",bg='#2C3E50' )
label1.pack(pady=10)

#Botões 
BtnI = Button(menu, text="Iniciar", command=lambda: btn_Click("iniciar"), bg='#27AE60', fg='#FFFFFF', width=15, height=2, relief='raised')
BtnI.pack(pady=10)
BtnI.pack(side=tk.LEFT, padx=20)  

BtnS = Button(menu, text="Parar", command=lambda: btn_Parar("parar"), bg='#C0392B', fg='#FFFFFF', width=15, height=2, relief='raised')
BtnS.pack(pady=10)
BtnS.pack(side=tk.LEFT, padx=5)  

btn_fechar = tk.Button(menu, text="X", command=btn_fechar, bg='#E74C3C', fg='#FFFFFF', font=('Helvetica', 10, 'bold'), bd=0)
btn_fechar.place(x=280, y=0, width=20, height=20)  


menu.mainloop()