from types import NoneType
import PySimpleGUI as sg
import funcoes
import webbrowser

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
            [sg.Text('Oprs.'), sg.Input(key= 'operadores', size= size_Input), sg.Text('Parada'), sg.Input(key= 'parada', size= size_Input), sg.Image(r'static/papaleguas.png')],
            [sg.Button('Confirmar'), sg.Button('Limpar', ), sg.Text(f'{"By: Leonardo Mantovani":>35}',enable_events= True, key= 'link', text_color=('blue'))],
            [sg.Output(size= (35, 15), key='__Output__', font= font_str)]
        ]
        # Janela
        self.window = sg.Window('OP Calculator v1.8', icon=r'static/papaleguas.ico').layout(layout)
    # Função da classe para a construção de todos os eventos de botões
    def start_program(self):
        while True:
            # Extrair os dados na tela
            # Cada event é um botão no programa, com um popup de ajuda
            try:
                event, self.values = self.window.Read()
                if event == sg.WIN_CLOSED or event == 'Sair':
                    break
                # Aqui será usada a função do DataScience para pegar os dados somente das produções que foram finalizadas com valor total
                if event == 'Cabeçalho':
                    try:
                        arquivo = sg.popup_get_file('Selecione o arquivo', 'Cabeçalho de ordem', icon=r'static/papaleguas.ico')
                        funcoes.organizar_cabecalho(arquivo)
                        print('Planilha concluída')
                    except:
                        print('Erro nas informações fornecidas')
                        pass
                # Aqui será usada a função do DataScience para pegar os dados de consumo dos componentes
                if event == 'Componentes':
                    try:
                        arquivo = sg.popup_get_file('Selecione o arquivo', 'Componentes', icon=r'static/papaleguas.ico')
                        funcoes.organizar_componentes(arquivo)
                        print('Planilha concluída')
                    except:
                        print('Erro nas informações fornecidas')
                        pass
                # Escolha o destino para salvar as planilhas
                if event == 'Definições':
                    destino = sg.popup_get_folder('Escolha o destino para salvar as planilhas', 'Salvar em', icon=r'static/papaleguas.ico')
                    if type(destino) != NoneType and len(destino) != 0:
                        funcoes.escolher_diretorio(destino)
                # Irá limpar o campo do Output aonde as informações são printadas, necessita o uso de melhor forma para o Output ainda !!
                if event == 'Limpar':
                    self.window.FindElement('__Output__').update('')
                # Link do GitHub do criador do programa
                if event == 'link':
                    webbrowser.open('https://github.com/LeonardoHMS')
                # Popup de ajuda de como preencher os campos do programa
                if event == 'Ajuda':
                    sg.popup(funcoes.text_popup(), title= 'Ajuda', icon=r'static/papaleguas.ico')
                # Será feita toda a conta matemática para gerar a quantidade em minutos do tempo de produção
                if event == 'Confirmar':
                    resultados = (funcoes.calcular_horario(self.values['inicio'], self.values['fim'],
                    self.values['operadores'], self.values['parada'].strip())) # Função para calcular o horário
                    # Irá exibir os resultados no Output
                    if resultados != False:
                        print(f'Tempo da Maquina: {resultados[0]:.0f} Minutos.')
                        print(f'Tempo de Mão humana: {resultados[1]:.0f} Minutos.')
                    # Caso as informações não estejam claramente informadas, será informado um erro ao usuário
                    else:
                        print(f'ERRO, Verifique as informações!')
            # Caso de um erro em algum dos campos de dados, uma mensagem será mostrada para o usuário de possiveis erros
            except:
                print('ERRO, verifique as informações fornecidas!')


if __name__ == '__main__':
    painel = program_painel()
    painel.start_program()