from classes import acessoSAP
from datetime import datetime
from types import NoneType
import PySimpleGUI as sg
import webbrowser
import funcoes

# Classe principal da criação do programa
class ProgramPainel:
    def __init__(self):
        _font_str = 'Arial, 12'
        _size_Input = (6,1)
        sg.change_look_and_feel('DarkGrey4')
        # Layout do programa
        _layout = [
            [sg.Menu([['Arquivos', ['Refugos', 'Ajuda', 'Sair']],
                ['Planilhas', ['Cabeçalho', 'Componentes', 'Tempo Operações', 'Definições']],
                ['Automático',['SAP - ENTE', 'Definir login', 'Teste Sem Planilha']]])
            ],
            [
                sg.Text('Dia Início'),
                sg.Input(key='d_inicio', size=(10,1), default_text=datetime.today().strftime('%d-%m-%Y')),
                sg.CalendarButton('C', format='%d-%m-%Y', month_names=funcoes.MES, day_abbreviations=funcoes.ABV_DIAS),
                sg.Text('Dia Fim'),
                sg.Input(key='d_fim', size=(10,1), default_text=datetime.today().strftime('%d-%m-%Y')),
                sg.CalendarButton('C', format='%d-%m-%Y', month_names=funcoes.MES, day_abbreviations=funcoes.ABV_DIAS)
            ],
            [
                sg.Text('Início'), 
                sg.Input(key = 'inicio', size=_size_Input, focus=True), 
                sg.Text('Fim'),
                sg.Input(key= 'fim', size=_size_Input,), 
                sg.Text(f'{"":>12}'), sg.Image(r'static/githublogo.png', key= 'link', tooltip='acessar', enable_events=True)
            ],
            [
                sg.Text('Oprs.'), 
                sg.Input(key= 'operadores', size=_size_Input), 
                sg.Text('Parada'), 
                sg.Input(key= 'parada', size=_size_Input), 
                sg.Text(f'{"By: LEONARDOHMS"}',enable_events=True, text_color=('black'), font='Arial, 10')
            ],
            [sg.Button('Confirmar', bind_return_key=True), sg.Button('Limpar')],
            [sg.Multiline(size=(35, 15), key='__Output__', font=_font_str, do_not_clear=False, disabled=True)]
        ]
        # Janela
        self.window = sg.Window('OP Calculator v2.52', icon=r'static/calculator.ico', return_keyboard_events=True, use_default_focus=False).layout(_layout)
        sg.cprint_set_output_destination(multiline_key='__Output__', window=self.window)
    # Função da classe para a construção de todos os eventos de botões
    def run(self):
        while True:
            # Extrair os dados na tela
            event, self.values = self.window.Read()
            if event == sg.WIN_CLOSED or event == 'Sair':
                self.window.close()
                break
            #try:
                # DataScience para pegar os dados somente das produções que foram finalizadas com valor total
            if event == 'Cabeçalho':
                arquivo = sg.popup_get_file('Selecione o arquivo', 'Cabeçalho de ordem', icon=r'static/calculator.ico')
                funcoes.organizar_cabecalho(arquivo, copy=True)
                sg.cprint('Planilha concluída')

            # DataScience para pegar os dados de consumo dos componentes
            if event == 'Componentes':
                arquivo = sg.popup_get_file('Selecione o arquivo', 'Componentes', icon=r'static/calculator.ico')
                funcoes.organizar_componentes(arquivo)
                sg.cprint('Planilha concluída')

            # DateScience para cálculo dos tempos de produção
            if event == 'Tempo Operações':
                arquivo = sg.popup_get_file('Selecione o arquivo', 'Tempo de operações', icon=r'static/calculator.ico')
                funcoes.organizar_tempos_prd(arquivo)
                sg.cprint('Planilha concluída')

            # Escolha o destino para salvar as planilhas
            if event == 'Definições':
                destino = sg.popup_get_folder('Salvar planilhas em', 'Local', icon=r'static/calculator.ico', default_path=funcoes.getDiretorio())
                if type(destino) != NoneType and len(destino) != 0:
                    funcoes.setDiretorio(destino)

            if event == 'SAP - ENTE':
                acessoSAP.printar()
                usuario, senha, acessosap = funcoes.getLoginSAP()
                salvar = funcoes.getDiretorio()
                Sap_cab = acessoSAP.SapGui(usuario, senha, acessosap)
                Sap_cab.conexaoSap('COOIS')
                Sap_cab.sapGetCabecalho()
                Sap_cab.gerarPlanilha(salvar, 'CABECALHO.XLSX')
                arquivo = f'{salvar}/CABECALHO.xlsx'
                funcoes.organizar_cabecalho(arquivo, text=True)
                Sap_cab.conexaoSap('COOIS')
                Sap_cab.sapGetComponentes(salvar, 'ordens.txt')
                Sap_cab.gerarPlanilha(salvar, 'COMPONENTES.XLSX')
                arquivo = f'{salvar}/COMPONENTES.xlsx'
                funcoes.organizar_componentes(arquivo)

            if event == 'Definir login':
                LoginSAP().RunApp()

            if event == 'Teste Sem Planilha':
                usuario, senha, acessoSAP = funcoes.getLoginSAP()
                destino = funcoes.getDiretorio()
                Sap_cab = acessoSAP.SapGui(usuario, senha, acessosap)
                Sap_cab.conexaoSap('COOIS')
                Sap_cab.GetCabecalhoSemPlanilha(destino)
                Sap_cab.conexaoSap('COOIS')
                Sap_cab.GetComponentesSemPlanilha(destino, 'ordens.txt')

            if event == 'Limpar':
                self.window.find_element('__Output__').update('')

            # Link do GitHub
            if event == 'link':
                webbrowser.open('https://github.com/LeonardoHMS')

            if event == 'Refugos':
                CalculoRefugo().RunApp()
            # Popup de ajuda de como preencher os campos do programa
            if event == 'Ajuda':
                sg.popup(funcoes.text_popup(), title= 'Ajuda', icon=r'static/calculator.ico')
            # Será feita toda a conta matemática para gerar a quantidade em minutos do tempo de produção
            if event == 'Confirmar':
                resultados = (funcoes.calcular_horario(self.values['inicio'], self.values['fim'],
                self.values['operadores'], self.values['parada'].strip(), self.values['d_inicio'], self.values['d_fim'])) # Função para calcular o horário
                if resultados != False:
                    sg.cprint(f'Tempo da Maquina: {resultados[0]:.0f} Minutos.')
                    sg.cprint(f'Tempo de Mão humana: {resultados[1]:.0f} Minutos.')
                else:
                    sg.cprint(f'ERRO, Verifique as informações!')
            # Caso de um erro em algum dos campos de dados, uma mensagem será mostrada para o usuário de possiveis erros
            #except:
            #    sg.cprint('Erro, verifique as informações fornecidas!')


class LoginSAP():
    def __init__(self):
        _, _, _acesso = funcoes.getLoginSAP()
        _size_Input = (14,1)
        sg.change_look_and_feel('DarkGrey4')
        _layout = [
            [sg.Text('Login'), sg.Input(key='login', size=_size_Input)],
            [sg.Text('Senha'), sg.Input(key='senha', size=_size_Input, password_char='*')],
            [sg.Text('acesso SAP'), sg.Input(key='acessosap', size=_size_Input, default_text=_acesso)],
            [sg.Button('Confirmar')]
        ]
        self.window = sg.Window('Login do SAP', icon=r'static/calculator.ico').layout(_layout)

    def RunApp(self):
        event, self.values = self.window.Read()
        if event == sg.WIN_CLOSED:
            self.window.close()
        if event == 'Confirmar':
            funcoes.setLoginSAP(self.values['login'], self.values['senha'], self.values['acessosap'])
            sg.cprint('Dados Fornecidos!!')
            self.window.close()


class CalculoRefugo():

    @staticmethod
    def replacePoint(valor):
        valor = valor.replace(',', '.')
        return float(valor)

    @staticmethod
    def resultadoRefugos(Trefugos, PesoT, total):
        Trefugos = CalculoRefugo.replacePoint(Trefugos)
        PesoT = CalculoRefugo.replacePoint(PesoT)
        total = CalculoRefugo.replacePoint(total)
        resultado = Trefugos * PesoT / total
        return resultado
        
    def __init__(self):
        self._size_Input = (7,1)
        sg.change_look_and_feel('DarkGrey4')
        _layout = [
            [sg.Text('Total   '), sg.Input(key='total', size=self._size_Input), sg.Text('Peso        '), sg.Input(key='PesoT', size=self._size_Input)],
            [sg.Text('Refugo'), sg.Input(key='Trefugos', size=self._size_Input), sg.Text('Resultado'), sg.Text(key='Prefugos')],
            [sg.Button('Confirmar', bind_return_key=True), sg.Text('Total Gasto'), sg.Text(key='peso_total')]
        ]
        self.window = sg.Window('Calculo de Refugos', icon=r'static/calculator.ico').layout(_layout)

    def RunApp(self):
        while True:
            event, self.values = self.window.Read()
            if event == sg.WIN_CLOSED:
                break
            if event == 'Confirmar':
                try:
                    resultado = CalculoRefugo.resultadoRefugos(self.values['Trefugos'], self.values['PesoT'], self.values['total'])
                    self.values['PesoT'] = CalculoRefugo.replacePoint(self.values['PesoT'])
                    peso_total = self.values['PesoT'] + resultado
                    self.window['peso_total'].update(str(f'{peso_total:.3f}').replace('.', ','))
                    self.window['Prefugos'].update(str(f'{resultado:.3f}').replace('.', ','))
                except:
                    sg.cprint('Preencha corretamente as informações')


if __name__ == '__main__':
    calculo = CalculoRefugo()
    calculo.RunApp()