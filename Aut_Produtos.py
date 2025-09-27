import openpyxl # type: ignore #Biblioteca que conecta o python com o excel
import pyperclip # type: ignore #Biblioteca que copia
import pyautogui # type: ignore #Biblioteca de automatização
from time import sleep #Biblioteca com temporizador

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
    pyautogui.click(172,337,duration=1)

    #Simula o teclado e nesse caso cola o que está copiado
    pyautogui.hotkey('ctrl','v')

    #Descrição produto
    desc = linha[1].value
    pyperclip.copy(desc)
    pyautogui.click(188,426,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(174,559,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Codigo produto
    cd_produto = linha[3].value
    pyperclip.copy(cd_produto)
    pyautogui.click(168,646,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(176,737,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(165,821,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(181,883,duration=1)
    #Temporizador
    sleep(3)

    #Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(172,369,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Quantidade do estoque
    qnt_estoque = linha[7].value
    pyperclip.copy(qnt_estoque)
    pyautogui.click(171,454,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Validade
    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(228,541,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Cor
    cor = linha[9].value
    pyperclip.copy(validade)
    pyautogui.click(212,628,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Tamanho
    pyautogui.click(198,713,duration=1)
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.click(176,741,duration=1)
        
    elif tamanho == 'Médio':
        pyautogui.click(204,762,duration=1)
       
    else:
        pyautogui.click(208,781,duration=1)
        
    #Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(184,795,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(185,864,duration=1)
    sleep(3)

    #Fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(175,405,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Origem
    origem = linha[13].value
    pyperclip.copy(origem)
    pyautogui.click(186,489,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Observação do produto
    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(188,586,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Codigo de barra
    cd_barra = linha[15].value
    pyperclip.copy(cd_barra)
    pyautogui.click(180,706,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Localização armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(189,792,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Confirmar
    pyautogui.click(177,852,duration=1)
    sleep(1)

    #Confirmar inclusão
    pyautogui.click(669,156,duration=1)

    #Repetir cadastro
    pyautogui.click(509,612,duration=1)

