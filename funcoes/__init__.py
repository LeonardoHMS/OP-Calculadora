from datetime import timedelta
import pandas as pd
import json
import pyperclip


def calcular_horario(inicio, fim, operadores, parada):
    '''Função para capturar o horário fornecido pelo usuário;
    --> Faz o calculo do tempo entre os dois horários com retorno em minutos
    - inicio: Horário inicial da produção
    - fim: Horário do término da produção
    - operadores: Quantidade de operadores para a produção
    - parada: Tempo de maquina parada sem produzir
    '''
    inicio_almoco = timedelta(hours=int(11), minutes=int(35), seconds=int(00))
    fim_almoco = timedelta(hours=int(13), minutes=int(5), seconds=int(00))
    if parada == '':
        parada = 0
    if operadores == '':
        operadores = 1
    try:
        for item in inicio:
            if item.isnumeric() == False:
                inicio = inicio.replace(item, '')
        for item in fim:
            if item.isnumeric() == False:
                fim = fim.replace(item, '')
        if len(inicio) != 4 or len(fim) != 4:
            return False
        inicio = timedelta(hours= int(inicio[:2]), minutes= int(inicio[2:]), seconds= 00)
        fim = timedelta(hours= int(fim[:2]), minutes= int(fim[2:]), seconds= 00)
        resultado = (fim - inicio).seconds
        if inicio <= inicio_almoco and fim >= fim_almoco:
            return ((resultado / 60 - 90) - int(parada), (resultado / 60 - 90) * int(operadores) - (int(parada) * int(operadores)))
        else:
            return ((resultado / 60) - int(parada), (resultado / 60) * int(operadores) - (int(parada) * int(operadores)))
    except:
        return False # Os retornos de falsos são para printar o erro no output do programa

# Função criada para não ficar muitas informações dentro da classe, pode ser usada para mais Textos futuramente !!!
def text_popup():
    texto = """ 
            Inicio: Inicio da produção
            Fim: Fim da produção
            Oprs.: Quantidade de operadores
            Parada: Tempo parado sem produzir
            """
    return texto

# Funções para Data Science para analisar somente as produções que foram finalizadas
def organizar_cabecalho(dir_cabecalho, copy=False):
    destino = getDiretorio()
    planilha = pd.read_excel(dir_cabecalho)
    planilha.insert(11, 'Qtde Falta', planilha['Quantidade da ordem (GMEIN)'] - planilha['Qtd.fornecida (GMEIN)'])
    remove_line = planilha[planilha['Qtde Falta'] > 0].index
    planilha = planilha.drop(remove_line)
    planilha.loc[planilha['Versão de produção'] == '0', 'Versão de produção'] = 0
    planilha.to_excel(f'{destino}/CabecalhoNew.xlsx', index=False)
    if copy:
        ordens = ''
        for ordem in planilha['Ordem']:
            ordens += f'{str(ordem)},'
            ordens = ordens.replace(',', '\n')
        pyperclip.copy(ordens)


def organizar_tempos_prd(dir_operacoes):
    destino = getDiretorio()
    tempos_df = pd.read_excel(dir_operacoes)
    tempos_df = tempos_df.drop(['Data início real de execução',
                                'Hora início real de execução',
                                'Data fim real da execução',
                                'Hora fim real da execução',
                                'Grupo','Tipo de roteiro',
                                'Duração processamen. (BEAZE)'], axis=1)

    tempos_df['HM'] = (tempos_df['Valor standard 2 (VGE02)'] / tempos_df['Quantidade básica (MEINH)']) * tempos_df['Qtd.boa total confirmada (MEINH)']
    tempos_df['Dif HM'] = tempos_df['Confirmação atividade 2 (ILE02)'] - tempos_df['HM']
    tempos_df['% de Dif HM'] = tempos_df['Dif HM'] / tempos_df['Confirmação atividade 2 (ILE02)']
    tempos_df.loc[tempos_df['HM']==0, 'HM'] = str('')
    tempos_df.loc[tempos_df['HM']=='', 'Dif HM'] = str('')
    tempos_df.loc[tempos_df['HM']==0, '% de Dif HM'] = str('')

    tempos_df['HH'] = (tempos_df['Valor standard 3 (VGE03)'] / tempos_df['Quantidade básica (MEINH)']) * tempos_df['Qtd.boa total confirmada (MEINH)']
    tempos_df['Dif HH'] = tempos_df['Atividade confirm.3 (ILE03)'] - tempos_df['HH']
    tempos_df['% de Dif HH'] = tempos_df['Dif HH'] / tempos_df['Atividade confirm.3 (ILE03)']
    tempos_df.loc[tempos_df['HH']==0, 'HH'] = str('')
    tempos_df.loc[tempos_df['HH']=='', 'Dif HH'] = str('')
    tempos_df.loc[tempos_df['HH']==0, '% de Dif HH'] = str('')

    tempos_df.to_excel(f'{destino}/TemposOperacoesNew.xlsx', index=False)


def organizar_componentes(dir_componentes):
    destino = getDiretorio()
    planilha = pd.read_excel(dir_componentes)
    planilha.insert(7, 'Qtde Falta', planilha['Qtd.necessária (EINHEIT)'] - planilha['Qtd.retirada (EINHEIT)'])
    planilha['OBS'] = ''
    planilha.to_excel(f'{destino}/ComponentesNew.xlsx', index=False)


def createDirectory():
    file_settings = r'static\Settings.json'
    try:
        with open(file_settings) as file:
            settings = json.load(file)
        return settings
    except FileNotFoundError:
        with open(file_settings, 'w') as file:
            informacoes = {}
            informacoes['Diretorio'] = 'Undefined'
            informacoes['Login'] = 'Undefined'
            informacoes['Senha'] = 'Undefined'
            informacoes['AcessoSAP'] = 'Undefined'
            json.dump(informacoes, file)
            settings = json.load(file)
        return settings


def setDiretorio(diretorio):
    settings = createDirectory()
    settings['Diretorio'] = diretorio
    with open(r'static\Settings.json', 'w') as file:
        json.dump(settings, file)


def getDiretorio():
    settings = createDirectory()
    return settings['Diretorio']


def setLoginSAP(usuario, senha, acesso):
    settings = createDirectory()
    settings['Login'] = usuario
    settings['Senha'] = senha
    settings['AcessoSAP'] = acesso
    with open(r'static\Settings.json', 'w') as file:
        json.dump(settings, file)
            

def getLoginSAP():
    settings = createDirectory()
    return [settings['Login'],
            settings['Senha'], 
            settings['AcessoSAP']
        ]