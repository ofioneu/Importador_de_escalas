import PySimpleGUI as sg
from viscofan.viscofan import Viscofan
from plastek.plastek import Plastek

class importadores:
    def __init__(self):
        sg.theme('BluePurple')   # Tema da tela
        #layout
        layout = [
            [sg.Text('Path do arquivo de entrada', size=(10,0)), sg.Input(key='input', size=(80, 0)), sg.FileBrowse()],
            [sg.Text('Path do arquivo de saída', size=(10,0)), sg.Input(key='output', size=(80, 0)), sg.FileBrowse()],
            [sg.Text('Exceção criada:', size=(0,0)), sg.Input(key='excecao')],
            [sg.Text('Marque a empresa:', font=(0,15))],
            [sg.Checkbox('Viscofan', key='viscofan')],
            [sg.Checkbox('Plastek', key='plastek')],
            [sg.Text('Log:', font=(0,15))],
            [sg.Output(text_color='red', size=(80,20))],
            [sg.Button(button_text='Processar'), sg.Button('Cancel')]
            ]
        #cria a tela
        self.window = sg.Window('IMPORTADOR DE ESCALAS').layout(layout)
    #função que inicia a tela  e a mantem rodando   
    def start(self):
        viscofan = Viscofan()
        plastek = Plastek()
        while True:
            #transfere os valores da tela  para a variavel value 
            self.Button, self.value=self.window.Read()
            
            path_ = self.value['input']
            output_ = self.value['output']
            excecao_ = self.value['excecao']
            
            if self.value['viscofan']==True:
                viscofan.execute(path_, output_, excecao_)
                if viscofan.execute(path_, output_, excecao_):
                    sg.Popup('Success!')

            if self.value['plastek']==True:
                plastek.execute(path_, output_, excecao_)
                if plastek.execute(path_, output_, excecao_):
                    sg.Popup('Success!')

            
            #Para o programa
            if self.Button in (None, 'Cancel'):   # if user closes window or clicks cancel
                break
        self.window.close()
try:
    a=importadores()
    a.start()
except:
    sg.Popup("An error occurred!")
