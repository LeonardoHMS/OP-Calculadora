from datetime import timedelta, datetime
import pandas as pd
import pyperclip
import json


ABV_DIAS = [
    'Dom',
    'Seg',
    'Ter',
    'Qua',
    'Qui',
    'Sex',
    'Sab'
]

MES = [
    'Janeiro',
    'Fevereiro',
    'Março',
    'Abril',
    'Maio',
    'Junho',
    'Julho',
    'Agosto',
    'Setembro',
    'Outubro',
    'Novembro',
    'Dezembro'
]

DIAS = [
    'Segunda-feira',
    'Terça-Feira',
    'Quarta-feira',
    'Quinta-feira',
    'Sexta-feira',
    'Sábado',
    'Domingo'
]

def somente_numeros(numeros):
    """ Função que deixa somentes números para formar um valor em timedelta
    --> Se a quantidade de números for diferente de 4, retorna False
        - numeros: Números usados para formar um horário
    """
    for item in numeros:
        if item.isnumeric() == False:
            numeros = numeros.replace(item, '')
    if len(numeros) != 4:
        return False
    numeros = timedelta(hours= int(numeros[:2]), minutes= int(numeros[2:]), seconds= 00)
    return numeros

def calcular_horario(inicio, fim, operadores, parada, d_inicio, d_fim):
    '''Função para capturar o horário fornecido pelo usuário;
    --> Faz o calculo do tempo entre os dois horários com retorno em minutos
    - inicio: Horário inicial da produção
    - fim: Horário do término da produção
    - operadores: Quantidade de operadores para a produção
    - parada: Tempo de maquina parada sem produzir
    - d_inicio: Dia inícial da produção
    - d_fim: Dia final da produção
    '''
    if '-' in d_inicio and d_fim:
        d_inicio = d_inicio.split('-')
        d_fim = d_fim.split('-')
    else:
        d_inicio = [d_inicio[:2], d_inicio[2:4], d_inicio[4:]]
        d_fim = [d_fim[:2], d_fim[2:4], d_fim[4:]]
    if parada == '':
        parada = 0
    if operadores == '':
        operadores = 1
    try:
        inicio = somente_numeros(inicio)
        d_inicio = datetime(
            day=int(d_inicio[0]),
            month=int(d_inicio[1]),
            year=int(d_inicio[2]),
        )
        d_inicio += inicio
        fim = somente_numeros(fim)
        d_fim = datetime(
            day=int(d_fim[0]),
            month=int(d_fim[1]),
            year=int(d_fim[2])
        )
        d_fim += fim
        resultado = 0
        while d_inicio < d_fim:
            indice_semana = d_inicio.weekday()
            dia_semana = DIAS[indice_semana]
            if d_inicio.hour == 11 and d_inicio.minute == 35:
                d_inicio += timedelta(minutes=90)
            elif d_inicio.hour == 17 and d_inicio.minute == 28:
                d_inicio += timedelta(minutes=822)
            else:
                d_inicio += timedelta(minutes=1)
                resultado += 1
            if dia_semana == 'Sábado':
                d_inicio += timedelta(days=2)
            elif dia_semana == 'Domingo':
                d_inicio += timedelta(days=1)
        return (resultado - int(parada), resultado * int(operadores) - (int(parada) * int(operadores)))
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
def organizar_cabecalho(dir_cabecalho, copy=False, text=False):
    destino = getDiretorio()
    planilha = pd.read_excel(dir_cabecalho)
    planilha.insert(11, 'Qtde Falta', planilha['Quantidade da ordem (GMEIN)'] - planilha['Qtd.fornecida (GMEIN)'])
    remove_line = planilha[planilha['Qtde Falta'] > 0].index
    planilha = planilha.drop(remove_line)
    planilha.loc[planilha['Versão de produção'] == '0', 'Versão de produção'] = 0
    if copy:
        ordens = planilha['Ordem'].to_string(index=False)
        pyperclip.copy(ordens)
    if text:
        with open(f'{destino}\ordens.txt', 'w+') as txt:
            txt.write(planilha['Ordem'].to_string(index=False))


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
    tempos_df.loc[tempos_df['HM']==0, 'HM'] = str('---')
    tempos_df.loc[tempos_df['HM']=='', 'Dif HM'] = str('---')
    tempos_df.loc[tempos_df['HM']==0, '% de Dif HM'] = str('---')

    tempos_df['HH'] = (tempos_df['Valor standard 3 (VGE03)'] / tempos_df['Quantidade básica (MEINH)']) * tempos_df['Qtd.boa total confirmada (MEINH)']
    tempos_df['Dif HH'] = tempos_df['Atividade confirm.3 (ILE03)'] - tempos_df['HH']
    tempos_df['% de Dif HH'] = tempos_df['Dif HH'] / tempos_df['Atividade confirm.3 (ILE03)']
    tempos_df.loc[tempos_df['HH']==0, 'HH'] = str('---')
    tempos_df.loc[tempos_df['HH']=='', 'Dif HH'] = str('---')
    tempos_df.loc[tempos_df['HH']==0, '% de Dif HH'] = str('---')

    tempos_df.to_excel(f'{destino}/TemposOperacoesNew.xlsx', index=False)


def organizar_componentes(dir_componentes):
    destino = getDiretorio()
    planilha = pd.read_excel(dir_componentes)
    planilha.insert(7, 'Qtde Falta', planilha['Qtd.necessária (EINHEIT)'] - planilha['Qtd.retirada (EINHEIT)'])
    planilha['Requirement date'] = planilha['Requirement date'].dt.strftime('%d/%m/%Y')
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