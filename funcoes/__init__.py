from datetime import timedelta
import pandas as pd

def calcular_horario(inicio, fim, operadores, com_almoco, sem_almoco, parada):
    '''Função para capturar o horário fornecido pelo usuário;
    --> Faz o calculo do tempo entre os dois horários com retorno em minutos
    - inicio: Horário inicial da produção
    - fim: Horário do término da produção
    - operadores: Quantidade de operadores para a produção
    - com_almoco: Subtrair o horário de almoço entre a produção
    - sem_almoco: Não incluir o horário de almoço entre a produção
    - parada: Tempo de maquina parada sem produzir
    '''
    if parada == '':
        parada = 0
    try:
        for item in inicio:
            if item.isnumeric() == False:
                inicio = inicio.replace(item, '')
        for item in fim:
            if item.isnumeric() == False:
                fim = fim.replace(item, '')
        if len(inicio) != 4 or len(fim) != 4:
            return False
        elif com_almoco and sem_almoco:
            return False
        inicio = timedelta(hours= int(inicio[:2]), minutes= int(inicio [2:]), seconds= 00)
        fim = timedelta(hours= int(fim[:2]), minutes= int(fim[2:]), seconds= 00)
        resultado = (fim - inicio).seconds
        if com_almoco:
            return ((resultado / 60 - 90) - int(parada), (resultado / 60 - 90) * int(operadores) - (int(parada) * int(operadores)))
        elif sem_almoco:
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
            Com almoço: Inclui o intervalo de almoço
            Sem almoco: Sem o intervalo de almoço
            """
    return texto

# Funções para Data Science para analisar somente as produções que foram finalizadas
def organizar_cabecalho(dir_cabecalho):
    destino = variavel_diretorio()
    planilha = pd.read_excel(f'{dir_cabecalho}')
    planilha.insert(11, 'Qtde Falta', planilha['Quantidade da ordem (GMEIN)'] - planilha['Qtd.fornecida (GMEIN)'])
    remove_line = planilha[planilha['Qtde Falta'] > 0].index
    planilha = planilha.drop(remove_line)
    planilha.to_excel(f'{destino}/CabecalhoNew.xlsx', index= False)


def organizar_componentes(dir_componentes):
    destino = variavel_diretorio()
    planilha = pd.read_excel(f'{dir_componentes}')
    planilha.insert(7, 'Qtde Falta', planilha['Qtd.necessária (EINHEIT)'] - planilha['Qtd.retirada (EINHEIT)'])
    planilha.to_excel(f'{destino}/ComponentesNew.xlsx', index= False)


def criar_sistema():
    try:
        sistema_arquivo = 'static\diretorios'
        arquivotxt = open(sistema_arquivo, 'r+')
    except FileNotFoundError:
        arquivotxt = open(sistema_arquivo, 'w+')
        arquivotxt.writelines('')
    return arquivotxt


def escolher_diretorio(diretorio):
    arquivotxt = criar_sistema()
    arquivotxt = open('static\diretorios', 'w')
    arquivotxt.write(f'{diretorio}')
    arquivotxt.close()


def variavel_diretorio():
    destino = open('static\diretorios', 'r+')
    for linha in destino:      
        return linha