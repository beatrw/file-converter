import os
import pandas as pd
import PySimpleGUI as sg

class TransformFIles():
    def __init__(self):
        self.path_inicio = None
        self.path_fim = None
    
    def _execute(self):
        self._info()
        self._get_path()
        self._transform()
        self._end()

    #Interface para informar o que o programa faz
    def _info(self):
        info = [
            [sg.Text("Olá! Aqui você pode converter arquivos .xls para .xlsx.")],
            [sg.Text("Para isso, basta informar o caminho da pasta de entrada e de saída na tela a seguir.")],
            [sg.Text("")],
            [sg.Text("Hello! Here you can convert .xls files to .xlsx.")],
            [sg.Text("To do so, just enter the input and output folder path on the following screen.")],
            [sg.Text("")],
            [sg.Button("Ok")]
            ]
        info = sg.Window('INFO', info) 
        info.read()
        info.close()

    #Interface para pegar o caminho das pastas
    def _get_path(self):
        #Interface
        path= [
            [sg.Text("Informe os caminhos das pastas de entrada e de saída")],
            [sg.Text("Enter the input and output folder paths")],
            [sg.Text("Pasta entrada | Input folder")],
            [sg.InputText(key="pasta_entrada")],
            [sg.Text("Pasta saída | Output folder")],
            [sg.InputText(key="pasta_saida")],
            [sg.Button("Ok")],
            ]
        path = sg.Window('Definição das pastas', path) 
        path_ = path.read()
        path.close()
        print(path_)
        self.path_inicio = path_[1].get('pasta_entrada').replace("\\", "/")
        self.path_fim = path_[1].get('pasta_saida').replace("\\", "/")

    def _transform(self):
        list_files = os.listdir(self.path_inicio)

        for file in list_files:
            df = pd.read_excel(self.path_inicio + "\\" + file)

            print(f"leu {file}")
            file = file.replace(".xls", ".xlsx")

            df.to_excel(self.path_fim + "\\" + file, index=False)
            print(f"transformou {file}")

    def _end(self):
        end = [
            [sg.Text("Arquivos convertidos com sucesso!")],
            [sg.Text("Files converted successfully!")],
            [sg.Button("Ok")]
            ]
        end = sg.Window('Fim', end) 
        end.read()
        end.close()
        
TransformFIles()._execute()