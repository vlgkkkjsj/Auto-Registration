import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1175,319,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1114,417,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1139,564,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1115,650,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1140,728,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1166,824,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1109,882,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1134,353,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1102,437,duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1185,527,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1122,613,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(1138,704,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1151,778,duration=1)
    elif tamanho == 'Medio':
        pyautogui.click(1122,769,duration=1)
    else :
        pyautogui.click(1114,798,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1156,783,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1121,849,duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1116,388,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1130,464,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value  
    pyperclip.copy(observacoes)
    pyautogui.click(1138,568,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1190,704,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1175,789,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(1132,852,duration=1)

    pyautogui.click(1598,596,duration=1)
    pyautogui.click(1437,606,duration=1)
