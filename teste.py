import pandas as pd

planilha = pd.read_excel(r"C:\Users\Leonardo Mantovani\Desktop\cabecalho.xlsx")
planilha.insert(11, 'Qtde Falta', planilha['Quantidade da ordem (GMEIN)'] - planilha['Qtd.fornecida (GMEIN)'])
remove_line = planilha[planilha['Qtde Falta'] > 0].index
planilha = planilha.drop(remove_line)
planilha.to_excel(r'C:\Users\Leonardo Mantovani\Desktop\novaplanilha.xlsx', index= False)
print(planilha.info())