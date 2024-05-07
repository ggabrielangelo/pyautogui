import openpyxl
import pyperclip 
import pyautogui
from time import sleep



workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente 
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(173,326, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(170,412, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(230,553, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(167,640, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(177,709, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(188,809, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(192,877, duration=1)
    sleep(3)


    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(189,358, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(176,449, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(171,533, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(191,610, duration=1)
    pyautogui.hotkey('ctrl', 'v')




    tamanho = linha[10].value
    pyautogui.click(183,699, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(261,729, duration=1)
    elif tamanho == "Médio":
        pyautogui.click(391,752, duration=1)
    else:
        pyautogui.click(210,769, duration=1)




    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(168,776, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(197,836, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(168,390, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(166,474, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(162,563, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(171,689, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(171,781, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(176,839, duration=1) 
    sleep(3)
    pyautogui.click(680,166,duration=1)
    sleep(2)
    pyautogui.click(539,610,duration=1)
    sleep(2)

