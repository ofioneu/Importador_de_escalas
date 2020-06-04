import PySimpleGUI as sg
from viscofan.viscofan import Viscofan

class importadores:
    def __init__(self):
        sg.theme('BluePurple')   # Tema da tela
        #layout
        layout = [
            [sg.Text('Path do arquivo de entrada', size=(10,0)), sg.Input(key='input')],
            [sg.Text('Path do arquivo de saída', size=(10,0)), sg.Input(key='output')],
            [sg.Text('Exceção criada:', size=(0,0)), sg.Input(key='excecao')],
            [sg.Text('Marque a empresa:', font=(0,15))],
            [sg.Checkbox('Viscofan', key='viscofan')],
            [sg.Checkbox('Plastek', key='plastek')],
            [sg.Button(button_text='Processar'), sg.Button('Cancel')]
            ]
        #cria a tela
        self.window = sg.Window('IMPORTADOR DE ESCALAS').layout(layout)
    #função que inicia a tela  e a mantem rodando   
    def start(self):
        viscofan = Viscofan()
        while True:
            #transfere os valores da tela  para a variavel value 
            self.Button, self.value=self.window.Read()
            if self.value['viscofan']==True:
                viscofan.execute(self.value['input'], self.value['output'], self.value['excecao'])
                

            
            #Para o programa
            if self.Button in (None, 'Cancel'):   # if user closes window or clicks cancel
                break
        self.window.close()

a=importadores()
a.start()