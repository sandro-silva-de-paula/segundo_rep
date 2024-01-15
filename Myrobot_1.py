
import openpyxl
from pathlib import Path
import pyperclip
import pyautogui
from time import sleep

CAMINHO = Path(__file__)
CAMINHO_ARQ = CAMINHO.parent / 'produtos_ficticios.xlsx'
workbook = openpyxl.load_workbook(CAMINHO_ARQ)
sheeet_produtos = workbook['Produtos']

for linha in sheeet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(128, 167, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(124, 258, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(124, 391, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(123, 476, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(125, 562, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(127, 652, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão próximo
    pyautogui.click(148, 702, duration=1)
    sleep(2)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(125, 194, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(124, 280, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(129, 367, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(127, 451, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(131, 537, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(131, 555, duration=1)
    elif tamanho == 'Medio':
        pyautogui.click(131, 573, duration=1)
    else:
        pyautogui.click(131, 591, duration=1)

    # pyautogui.hotkey('ctrl', 'v')

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(129, 625, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão próximo
    pyautogui.click(146, 683, duration=1)
    sleep(2)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(134, 215, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(136, 302, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    obs_ = linha[14].value
    pyperclip.copy(obs_)
    pyautogui.click(142, 388, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(140, 516, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizador_armazem = linha[16].value
    pyperclip.copy(localizador_armazem)
    pyautogui.click(139, 607, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão concluir
    pyautogui.click(162, 666, duration=1)
    sleep(2)

    # Botão produto salvo
    pyautogui.click(855, 169, duration=1)
    sleep(2)

    # Botão proximo produto
    pyautogui.click(675, 447, duration=1)
    sleep(2)
