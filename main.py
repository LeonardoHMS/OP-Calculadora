# --------------------------------------------------------------------------------------
# Programa criado Leonardo Mantovani, github.com/LeonardoHMS
# --------------------------------------------------------------------------------------
# Criado como treino de aprendizagem da linguagem python, que se tornou bem útil !!!
# --------------------------------------------------------------------------------------
# Utilizo esse programa em meu local de trabalho
# --------------------------------------------------------------------------------------
from classes.frontend_GUI import ProgramPainel
import funcoes

# Load Settings
try:
    funcoes.createDirectory()
except:
    pass

iniciar = ProgramPainel()
iniciar.run()
