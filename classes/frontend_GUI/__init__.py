from types import NoneType
import PySimpleGUI as sg
import funcoes
import webbrowser
from classes import acessoSAP

# Classe principal da criação do programa
class ProgramPainel:
    def __init__(self):
        font_str = 'Arial, 12'
        size_Input = (6,1)
        sg.change_look_and_feel('DarkGrey4')
        # Layout do programa
        layout = [# Menu da parte de cima do programa
            [sg.Menu([['Arquivos', ['Ajuda', 'Sair']],
                ['Planilhas', ['Cabeçalho', 'Componentes', 'Tempo Operações', 'Definições']],
                ['Automático',['SAP - Cabeçalho', 'Definir login']]])],# Menu da parte de cima do programa
            [
                sg.Text('Inicio'), 
                sg.Input(key = 'inicio', size= size_Input), 
                sg.Text('Fim'),
                sg.Input(key= 'fim', size= size_Input,), 
                sg.Text(f'{"":>12}'), sg.Image(r'static/githublogo.png', key= 'link', tooltip='acessar', enable_events=True)
            ],
            [
                sg.Text('Oprs.'), 
                sg.Input(key= 'operadores', size= size_Input), 
                sg.Text('Parada'), 
                sg.Input(key= 'parada', size= size_Input), 
                sg.Text(f'{"By: LEONARDOHMS"}',enable_events=True, text_color=('black'), font='Arial, 10')
            ],
            [sg.Button('Confirmar'), sg.Button('Limpar')],
            [sg.Output(size= (35, 15), key='__Output__', font= font_str)]
        ]
        # Janela
        self.window = sg.Window('OP Calculator v2.01', icon=r'static/calculator.ico').layout(layout)
    # Função da classe para a construção de todos os eventos de botões
    def startProgram(self):
        while True:
            # Extrair os dados na tela
            try:
                event, self.values = self.window.Read()
                if event == sg.WIN_CLOSED or event == 'Sair':
                    break
                # DataScience para pegar os dados somente das produções que foram finalizadas com valor total
                if event == 'Cabeçalho':
                    arquivo = sg.popup_get_file('Selecione o arquivo', 'Cabeçalho de ordem', icon=r'static/calculator.ico')
                    funcoes.organizar_cabecalho(arquivo)
                    print('Planilha concluída')

                # DataScience para pegar os dados de consumo dos componentes
                if event == 'Componentes':
                    arquivo = sg.popup_get_file('Selecione o arquivo', 'Componentes', icon=r'static/calculator.ico')
                    funcoes.organizar_componentes(arquivo)
                    print('Planilha concluída')

                # DateScience para cálculo dos tempos de produção
                if event == 'Tempo Operações':
                    arquivo = sg.popup_get_file('Selecione o arquivo', 'Tempo de operações', icon=r'static/calculator.ico')
                    funcoes.organizar_tempos_prd(arquivo)
                    print('Planilha concluída')

                # Escolha o destino para salvar as planilhas
                if event == 'Definições':
                    destino = sg.popup_get_folder('Salvar planilhas em', 'Local', icon=r'static/calculator.ico', default_path=funcoes.getDiretorio())
                    if type(destino) != NoneType and len(destino) != 0:
                        funcoes.setDiretorio(destino)

                if event == 'SAP - Cabeçalho':
                    usuario, senha, acessosap = funcoes.getLoginSAP()
                    Sap_cab = acessoSAP.SapGui(usuario, senha, acessosap)
                    Sap_cab.conexaoSap('COOIS')
                    salvar = funcoes.getDiretorio()
                    Sap_cab.sapCooisXlsx(salvar)
                    arquivo = f'{salvar}/EXPORT.xlsx'
                    funcoes.organizar_cabecalho(arquivo)

                if event == 'Definir login':
                    LoginSAP()

                # Irá limpar o campo do Output aonde as informações são printadas, necessita o uso de melhor forma para o Output ainda !!
                if event == 'Limpar':
                    self.window.FindElement('__Output__').update('')

                # Link do GitHub do criador do programa
                if event == 'link':
                    webbrowser.open('https://github.com/LeonardoHMS')

                # Popup de ajuda de como preencher os campos do programa
                if event == 'Ajuda':
                    sg.popup(funcoes.text_popup(), title= 'Ajuda', icon=r'static/calculator.ico')

                # Será feita toda a conta matemática para gerar a quantidade em minutos do tempo de produção
                if event == 'Confirmar':
                    resultados = (funcoes.calcular_horario(self.values['inicio'], self.values['fim'],
                    self.values['operadores'], self.values['parada'].strip())) # Função para calcular o horário
                    if resultados != False:
                        print(f'Tempo da Maquina: {resultados[0]:.0f} Minutos.')
                        print(f'Tempo de Mão humana: {resultados[1]:.0f} Minutos.')
                    else:
                        print(f'ERRO, Verifique as informações!')
            # Caso de um erro em algum dos campos de dados, uma mensagem será mostrada para o usuário de possiveis erros
            except:
                print('ERRO, verifique as informações fornecidas!')


class LoginSAP():
    def __init__(self):
        _, _, acesso = funcoes.getLoginSAP()
        size_Input = (14,1)
        sg.change_look_and_feel('DarkGrey4')
        layout = [
            [sg.Text('Login '), sg.Input(key='login', size=size_Input)],
            [sg.Text('Senha'), sg.Input(key='senha', size=size_Input, password_char='*')],
            [sg.Text('acesso SAP'), sg.Input(key='acessosap', size=size_Input, default_text=acesso)],
            [sg.Button('Confirmar')]
        ]
        self.window = sg.Window('Login do SAP', icon=r'static/calculator.ico').layout(layout)
        event, self.values = self.window.Read()
        if event == 'Confirmar':
            funcoes.setLoginSAP(self.values['login'], self.values['senha'], self.values['acessosap'])
            print('Dados Fornecidos!!')
            self.window.close()


if __name__ == '__main__':
    painel = ProgramPainel()
    painel.startProgram()