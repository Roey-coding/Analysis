import PySimpleGUI as sg
from datetime import datetime, timedelta
from openpyxl import Workbook
from os import path
import openpyxl
import re

layout = [[sg.Text('We would like to create a root sheet that will contain all of your links in one master sheet, In order to do so while keeping it secured we need you to set a password to it so no one will be able to change it but you.')], [sg.Text('password: '), sg.InputText(password_char='*')], [sg.OK()]]
window = sg.Window("analysis", layout)

while True:
	event, values = window.read()
	
	if(event in (sg.WIN_CLOSED,  'Exit')):
		print(values)
		break
		
window.close()