import PySimpleGUI as sg
import funcoes
import webbrowser

diretorio = 'static\local_arquivo'
local_planilha = open(diretorio, 'r+')
dir_cabecalho = dir_new_cabecalho = 0
for cont, line in enumerate(local_planilha):
    if cont == 0:
        dir_cabecalho = line[:-1]
    else:
        dir_new_cabecalho = line

class program_painel:
    def __init__(self):
        font_str = 'Arial, 12'
        size_Input = (8,1)
        sg.change_look_and_feel('DarkGrey4')
        # Layout do programa
        layout = [
            [sg.Menu([['Arquivos', ['Ajuda', 'Sair']], ['Planilhas', ['Cabeçalho', 'Componentes', 'Definições']]],
            )],
            [sg.Text('Inicio'), sg.Input(key = 'inicio', size= size_Input), sg.Text('Fim'),
            sg.Input(key= 'fim', size= size_Input,)],
            [sg.Text('Oprs.'), sg.Input(key= 'operadores', size= size_Input), sg.Text('Parada'), sg.Input(key= 'parada', size= size_Input)],
            [sg.Checkbox('Com almoço', key= 'com_almoco'), sg.Checkbox(f'{"Sem almoço":<20}', key = 'sem_almoco'), sg.Image(r'static/papaleguas.png')],
            [sg.Button('Confirmar'), sg.Button('Limpar'), sg.Text(f'{"By: Leonardo Mantovani":>35}',enable_events= True, key= 'link')],
            [sg.Output(size= (35, 15), key='__Output__', font= font_str)],
        ]

        # Janela
        self.window = sg.Window('OP Calculator v1.0', icon=r'static/papaleguas.ico').layout(layout)

    def start_program(self):
        while True:
            # Extrair os dados na tela
            # Cada event é um botão no programa, com um popup de ajuda
            try:
                event, self.values = self.window.Read()
                if event == sg.WIN_CLOSED or event == 'Sair':
                    break
                if event == 'Cabeçalho':
                    funcoes.organizar_cabecalho(dir_cabecalho, dir_new_cabecalho)
                if event == 'Limpar':
                    self.window.FindElement('__Output__').update('')
                if event == 'link':
                    webbrowser.open('https://github.com/LeonardoHMS')
                if event == 'Ajuda':
                    sg.popup(funcoes.text_popup(), title= 'Ajuda', icon=r'static/papaleguas.ico')
                if event == 'Confirmar':
                    inicio = self.values['inicio']
                    fim = self.values['fim']
                    operadores = self.values['operadores']
                    parada = self.values['parada']
                    if parada == '':
                        parada = '0'
                    com_almoco = self.values['com_almoco']
                    sem_almoco = self.values['sem_almoco']
                    informacoes = (inicio, fim, operadores, com_almoco, sem_almoco, parada)
                    resultados = (funcoes.calcular_horario(informacoes))
                    if resultados != False:
                        print(f'Tempo da Maquina: {resultados[0]:.0f} Minutos.')
                        print(f'Tempo de Mão humana: {resultados[1]:.0f} Minutos.')
                    else:
                        print(f'ERRO, Verifique as informações!')
            except:
                print('ERRO, preencha todos os campos !')

