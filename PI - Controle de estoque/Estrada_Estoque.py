from tkinter import font
from tkinter.font import BOLD
import PySimpleGUI as sg
from numpy import flexible
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
import pathlib


sg.theme("DarkGrey2")


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
            button_color=("black", "green"),
        ),
        sg.Button("Limpar", size=(15, 1), font=("Roboto", 12, BOLD)),
        sg.Exit(
            "Sair",
            size=(15, 1),
            font=("Roboto", 12, BOLD),
            button_color=("black", "red"),
        ),
    ],
]


window = sg.Window(
    "Entrada de estoque",
    layout,
    size=(650, 150),
    element_justification="center",
)


def clear_input():
    for key in values:
        window[key].update(" ")
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Sair":
        break

    if event == "Limpar":
        clear_input()

    if event == "Adicionar":
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup("Item adicionado com sucesso!")
        clear_input()

window.close()
