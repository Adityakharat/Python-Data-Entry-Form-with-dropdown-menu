from typing import Sized, Text
from flask import Flask
import PySimpleGUI as sg
import pandas as pd
from PySimpleGUI.PySimpleGUI import Checkbox, Input, Spin

sg.theme('DarkTeal9')

EXCEL_FILE = 'data entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout =[
    [sg.Text('Please fill out the forecast form:')],
    [sg.CalendarButton('Date', target=(1,0), key='Date')],
    [sg.Text('Name', size = (15, 1)),sg.InputText(key='Name')],
    [sg.Text('Product', size = (15, 1)), sg.Combo(['Apple Royal Gala 90 count', 'Apple Royal Gala 100 count 20/22kg ', 'Apple Royal Gala 125 count', 'Apple Royal Gala 150 count', 'Apple Red Delicious 90 count', 'Apple Red Delicious 90 count', 'Apple Red Delicious 90 count', 'Apple Red Delicious 113 count', 'Apple Red Delicious 100 count', 'Apple Red Delicious 125 count', 'Apple Red Delicious 150 count', 'Mr. Apple 130 count', 'Indian Apple 100 count', 'Indian Apple 150 count', 'Indian Apple 175 count', 'Indian Apple Pittu 310 count', 'Indian Pear', 'Kiwi', 'Kiwi Gold', 'Grapes Red China', 'Indian orange', 'Orange', 'Dragon Fruit', 'Guava', 'Sweet Lime', 'Papaya', 'Pomegranate', 'Plum', 'Sapota', 'Elichi Banana'], key='Product')],
    [sg.Text('Target Price', size = (15, 1)),sg.InputText(key='Target price')],
    [sg.Text('UOM', size = (15, 1)),
                            sg.Checkbox('Kg', key = 'Kg'),
                            sg.Checkbox('Box', key = 'Box'),
                            sg.Checkbox('Crate', key = 'Crate'),
                            sg.Checkbox('Piece', key = 'Piece'),
                            sg.Checkbox('Punnet', key = 'Punnet')],
    [sg.Text('Quantity', size = (15,1)), sg.Spin([i for i in range (0,100)],
                                                  initial_value=0, key= 'Quantity')],                       

    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]   

window = sg.Window('Forecast entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        clear_input()
        sg.popup('Data saved!')
window.close()