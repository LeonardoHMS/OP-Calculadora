import PySimpleGUI as sg
import funcoes
import webbrowser

# Retirando as informaçoes do arquivo de texto para as variaveis assumirem o diretório para as planilhas
diretorio = 'static\local_arquivo'
local_planilha = open(diretorio, 'r+')
dir_cabecalho = dir_new_cabecalho = dir_componentes = dir_new_componentes = 0
for cont, line in enumerate(local_planilha):
    if cont == 0:
        dir_cabecalho = line[:-1]
    elif cont == 1:
        dir_new_cabecalho = line[:-1]
    elif cont == 2 :
        dir_componentes = line[:-1]
    else:
        dir_new_componentes = line
# Classe principal da criação do programa
class program_painel:
    def __init__(self):
        font_str = 'Arial, 12'
        size_Input = (8,1)
        sg.change_look_and_feel('DarkGrey4')
        # Layout do programa
        layout = [
            [sg.Menu([['Arquivos', ['Ajuda', 'Sair']], ['Planilhas', ['Cabeçalho', 'Componentes', 'Definições']]], # Menu da parte de cima do programa
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
# Função da classe para a construção de todos os eventos de botões
    def start_program(self):
        while True:
            # Extrair os dados na tela
            # Cada event é um botão no programa, com um popup de ajuda
            try:
                event, self.values = self.window.Read()
                if event == sg.WIN_CLOSED or event == 'Sair':
                    break
                if event == 'Cabeçalho': # Aqui será usada a função do DataScience para pegar os dados somente das produções que foram finalizadas com valor total
                    funcoes.organizar_cabecalho(dir_cabecalho, dir_new_cabecalho)
                if event == 'Limpar': # Irá limpar o campo do Output aonde as informações são printadas, necessita o uso de melhor forma para o Output ainda !!
                    self.window.FindElement('__Output__').update('')
                if event == 'link':
                    webbrowser.open('https://github.com/LeonardoHMS') # Link do GitHub do criador do programa
                if event == 'Ajuda':
                    sg.popup(funcoes.text_popup(), title= 'Ajuda', icon=r'static/papaleguas.ico') # Popup de ajuda de como preencher os campos do programa
                if event == 'Confirmar': # Será feita toda a conta matemática para gerar a quantidade em minutos do tempo de produção
                    inicio = self.values['inicio']
                    fim = self.values['fim']
                    operadores = self.values['operadores']
                    parada = self.values['parada']
                    if parada == '': # Caso não tenha parada, o usuário não precisa informar o valor zero
                        parada = '0'
                    com_almoco = self.values['com_almoco']
                    sem_almoco = self.values['sem_almoco']
                    informacoes = (inicio, fim, operadores, com_almoco, sem_almoco, parada) # uma Tupla com as informações para uso da função do calculo
                    resultados = (funcoes.calcular_horario(informacoes)) # Função para calcular o horário
                    if resultados != False: # Irá exibir os resultados no Output
                        print(f'Tempo da Maquina: {resultados[0]:.0f} Minutos.')
                        print(f'Tempo de Mão humana: {resultados[1]:.0f} Minutos.')
                    else: # Caso as informações não estejam claramente informadas, será informado um erro ao usuário
                        print(f'ERRO, Verifique as informações!')
            except: # Caso de um erro em algum dos campos de dados, uma mensagem será mostrada para o usuário de possiveis erros
                print('ERRO, preencha todos os campos !')

