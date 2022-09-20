from tkinter import Button, font
from tkinter.font import BOLD
import PySimpleGUI as sg
from numpy import flexible
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
import pathlib


sg.theme("Reddit")


# Caminho do excel (pos import pandas as pd)
EXCEL_FILE = "Estoque.xlsx"

arquivo = pathlib.Path(EXCEL_FILE)

if arquivo.exists():
    pass
else:
    sg.popup("Criando pasta de estoque... tente abrir novamente")
    EXCEL_FILE = Workbook()
    sheet = EXCEL_FILE.active
    sheet["A1"] = "-FUNCIONARIO-"
    sheet["B1"] = "-ADICIONADO-"
    sheet["C1"] = "-VENCE-"
    sheet["D1"] = "-QUANTIDADE-"
    sheet["E1"] = "-CODIGO-"
    sheet["F1"] = "-PRODUTO-"
    sheet["G1"] = "-PRECO-"

    EXCEL_FILE.save("Estoque.xlsx")


df = pd.read_excel(EXCEL_FILE)


funcionarios = ["Carlos", "Lucia", "Leonardo", "Luana"]


def janela_login():
    sg.theme("Reddit")
    layout = [
        [
            [sg.Text("Usuario"), sg.Input(key="-USUARIO-", size=(20, 1))],
            [sg.Button("Continuar", key="-BOTAOCONTINUAR-"), sg.Button("Sair")],
        ]
    ]

    return sg.Window("Login", layout=layout, finalize=True)


def janela_principal():
    sg.theme("Reddit")
    layout = [
        [sg.Text("Bem vindo ao sistema de estoque", font=("Arial", 20))],
        [sg.Text("Escolha uma opção abaixo:", font=("Arial", 15))],
        [
            sg.Button("Adicionar produto", size=(15, 1), key="-BOTAOADICIONAR-"),
            sg.Button("Remover produto", size=(15, 1), key="-BOTAOREMOVER-"),
        ],
        [
            sg.Button("Visualizar estoque", size=(20, 1), key="-BOTAOVISUALIZAR-"),
            sg.Button("Verificar vencimentos", size=(20, 1), key="-BOTAOVERIFICAR-"),
        ],
        [sg.Button("Voltar", size=(10, 1)), sg.Button("Sair", size=(10, 1))],
    ]
    return sg.Window(
        "Estoque", layout=layout, element_justification="center", finalize=True
    )


def Adicionar_Produtos():
    sg.theme("Reddit")
    layout = [
        [
            sg.Text("Funcionario:", size=(15, 1), font=("Roboto", 12)),
            sg.Push(),
            sg.InputCombo(funcionarios, size=(20, 1), key="-FUNCIONARIO-"),
            sg.Push(),
            sg.CalendarButton(
                "Data",
                target="-ADICIONADO-",
                format="%d/%m/%y",
                font=("Roboto", 10, BOLD),
            ),
            sg.InputText(size=(10, 1), key="-ADICIONADO-"),
            sg.Push(),
            sg.CalendarButton(
                "Vencimento",
                target="-VENCE-",
                format="%d/%m/%y",
                font=("Roboto", 10, BOLD),
            ),
            sg.InputText(size=(8, 1), key="-VENCE-"),
        ],
        [
            sg.Text("Codigo de barra:", size=(13, 1), font=("Roboto", 12)),
            sg.InputText(size=(26, 1), key="-CODIGO-", font=("Roboto", 12)),
            sg.Push(),
            sg.Text("Quantidade:", size=(9, 1), font=("Roboto", 12)),
            sg.InputText(size=(11, 1), key="-QUANTIDADE-", font=("Roboto", 12)),
        ],
        [
            sg.Text("Nome do produto:", size=(14, 1), font=("Roboto", 12)),
            sg.InputText(size=(25, 1), key="-PRODUTO-", font=("Roboto", 12)),
            sg.Push(),
            sg.Text("Preco: ", size=(5, 1), font=("Roboto", 12)),
            sg.InputText(size=(11, 1), key="-PRECO-", font=("Roboto", 12)),
        ],
        [sg.VPush()],
        [
            sg.Submit(
                "Adicionar",
                size=(15, 1),
                font=("Roboto", 12, BOLD),
                button_color=("white", "green"),
                key="-BOTAOADICIONARPRODUTO-",
            ),
            sg.Button(
                "Limpar", size=(15, 1), font=("Roboto", 12, BOLD), key="-LIMPAR-"
            ),
            sg.Exit(
                "Sair",
                size=(15, 1),
                font=("Roboto", 12, BOLD),
                button_color=("white", "red"),
            ),
        ],
    ]
    return sg.Window("Adicionar produtos", layout=layout, finalize=True)


def clear_input():
    for key in values:
        window[key].update(" ")
    return None


usuariosLogin = ["1", "Davi", "Carlos", "Othavio"]

janela1, janela2, janela3 = janela_login(), None, None

while True:
    window, event, values = sg.read_all_windows()

    if window == janela1 and event == sg.WIN_CLOSED or event == "Sair":
        break
    if window == janela1 and event == "-BOTAOCONTINUAR-":
        if values["-USUARIO-"] in usuariosLogin:
            janela1.hide()
            janela2 = janela_principal()
        else:
            sg.popup("Usuario invalido")
    # Aqui abriu a janela principal com os menus
    if window == janela2 and event == sg.WIN_CLOSED or event == "Sair":
        break

    # No menu da janela principal, se clicar em adicionar produto, abre a janela de adicionar produtos
    if window == janela2 and event == "-BOTAOADICIONAR-":
        janela2.hide()
        janela3 = Adicionar_Produtos()
    if window == janela3 and event == sg.WIN_CLOSED or event == "Sair":
        break
    if (
        event == janela3 and event == "-LIMPAR-"
    ):  # Se clicar em limpar, limpa os campos da janela de adicionar produtos
        clear_input()
    if (
        event == janela3 and event == "-BOTAOADICIONARPRODUTO-"
    ):  # Adicionar produto no excel
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup("Item adicionado com sucesso!")
        clear_input()

janela_principal.close()
