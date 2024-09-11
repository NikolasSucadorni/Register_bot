import openpyxl
import pyautogui

aba = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = aba['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(695,283,duration=1.3)
    pyautogui.write(linha[0].value)
    pyautogui.click(695,310,duration=1.3)
    pyautogui.write(linha[1].value) 
    pyautogui.click(689,336,duration=1.3)
    pyautogui.write(str(linha[2].value))
    pyautogui.click(824,364,duration=1.3)
    pyautogui.write(linha[3].value)
    pyautogui.click(620,391,duration=1.3)
    pyautogui.click(656,429,duration=1.3)