from datetime import timedelta
import pandas as pd

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
        if inicio < inicio_almoco and fim > fim_almoco:
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
def organizar_cabecalho(dir_cabecalho):
    destino = variavel_diretorio()
    planilha = pd.read_excel(f'{dir_cabecalho}')
    planilha.insert(11, 'Qtde Falta', planilha['Quantidade da ordem (GMEIN)'] - planilha['Qtd.fornecida (GMEIN)'])
    remove_line = planilha[planilha['Qtde Falta'] > 0].index
    planilha = planilha.drop(remove_line)
    planilha.loc[planilha['Versão de produção'] == '0', 'Versão de produção'] = 0
    planilha.to_excel(f'{destino}/CabecalhoNew.xlsx', index= False)


def organizar_componentes(dir_componentes):
    destino = variavel_diretorio()
    planilha = pd.read_excel(f'{dir_componentes}')
    planilha.insert(7, 'Qtde Falta', planilha['Qtd.necessária (EINHEIT)'] - planilha['Qtd.retirada (EINHEIT)'])
    planilha['OBS'] = ''
    planilha.to_excel(f'{destino}/ComponentesNew.xlsx', index= False)


def criar_sistema():
    try:
        sistema_arquivo = 'static\diretorios'
        arquivotxt = open(sistema_arquivo, 'r+')
        arquivotxt.close()
    except FileNotFoundError:
        arquivotxt = open(sistema_arquivo, 'w+')
        arquivotxt.writelines('Diretorio: \nLogin: \nSenha: \nAcessoSAP: ')
        arquivotxt.close()
    return arquivotxt


def escolher_diretorio(diretorio):
    arquivotxt = criar_sistema()
    with open('static\diretorios', 'r') as local:
        arquivotxt = local.readlines()
    with open('static\diretorios', 'w') as local:
        for line in arquivotxt:
            if arquivotxt.index(line) == 0:
                local.write(f'Diretorio: {diretorio}\n')
            else:
                local.write(line)


def variavel_diretorio():
    destino = criar_sistema()
    destino = open('static\diretorios', 'r')
    diretorio = destino.readline()
    return diretorio[11:-1]


def escolher_loginSAP(usuario, senha, acesso):
    with open('static\diretorios', 'r') as local:
        arquivotxt = local.readlines()
    with open('static\diretorios', 'w') as local:
        for line in arquivotxt:
            if arquivotxt.index(line) == 0:
                local.write(line)
            if arquivotxt.index(line) == 1:
                local.write(f'Login: {usuario}\n')
            if arquivotxt.index(line) == 2:
                local.write(f'Senha: {senha}\n')
            if arquivotxt.index(line) == 3:
                local.write(f'AcessoSAP: {acesso}\n')
            

def variavel_loginSAP():
    destino = open('static\diretorios', 'r')
    for cont, lines in enumerate(destino.readlines()):
        if cont == 1:
            usuario = lines[7:-1]
        if cont == 2:
            senha = lines[7:-1]
        if cont == 3:
            acessosap = lines[11:-1]
    return usuario, senha, acessosap