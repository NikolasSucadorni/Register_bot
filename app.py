import openpyxl
import pyautogui

aba = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = aba['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(655,360,duration=1.3)
    pyautogui.write(linha[0].value)
    pyautogui.click(660,383,duration=1.3)
    pyautogui.write(linha[1].value) 
    pyautogui.click(661,411,duration=1.3)
    pyautogui.write(str(linha[2].value))
    pyautogui.click(744,436,duration=1.3)
    pyautogui.write(linha[3].value)
    pyautogui.click(595,463,duration=1.3)
    pyautogui.click(653,425,duration=1.3)