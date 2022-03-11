from datetime import timedelta

def calcular_horario(lista):
    '''Função para capturar o horário fornecido pelo usuário;
    Fazer o calculo do tempo entre os dois horários com retorno em minutos
    lista -> Recebe uma lista com todas as informações necessárias para o cálculo
    inicio: Horário inicial da produção
    fim: Horário do término da produção
    operadores: Quantidade de operadores para a produção
    com_almoco: Subtrair o horário de almoço entre a produção
    sem_almoco: Não incluir o horário de almoço entre a produção
    parada: Tempo de maquina parada sem produzir
    '''
    try:
        inicio, fim, operadores, com_almoco, sem_almoco, parada = lista
        for item in inicio:
            if item.isnumeric() == False:
                inicio = inicio.replace(item, '')
        for item in fim:
            if item.isnumeric() == False:
                fim = fim.replace(item, '')
        if len(inicio) > 4 or len(fim) > 4:
            return False
        if inicio == '' or fim == '':
            return False
        elif com_almoco and sem_almoco:
            return False
        inicio_horas = int(inicio[:2])
        inicio_minutos = int(inicio [2:])
        fim_horas = int(fim[:2])
        fim_minutos = int(fim[2:])
        inicio = timedelta(hours= inicio_horas, minutes= inicio_minutos, seconds= 00)
        fim = timedelta(hours= fim_horas, minutes= fim_minutos, seconds= 00)
        resultado = fim - inicio
        if com_almoco:
            return [(resultado.seconds / 60 - 90) - int(parada), (resultado.seconds / 60 - 90) * int(operadores) - (int(parada) * int(operadores))]
        elif sem_almoco:
            return [(resultado.seconds / 60) - int(parada), (resultado.seconds / 60) * int(operadores) - (int(parada) * int(operadores))]
    except:
        return False