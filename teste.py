diretorio = 'static\local_arquivo'
local_planilha = open(diretorio, 'r+')
dir_cabecalho = dir_new_cabecalho = 0
for cont, line in enumerate(local_planilha):
    if cont == 0:
        dir_cabecalho = line[:-1]
    else:
        dir_new_cabecalho = line
print(dir_cabecalho)
print('----')
print(dir_new_cabecalho)