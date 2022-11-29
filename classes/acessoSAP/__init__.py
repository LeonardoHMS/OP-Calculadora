# -------------------------
# Classe para acesso ao SAP
# Abertura e Login automático do FrontEnd SAP
# Integração com SAP GUI Scripting API
# Leonardo Mantovani github.com/LeonardoHMS
# -------------------------
import win32com.client as win32
import pandas as pd
import subprocess
import time


class SapGui(object):
    """
        Cria um objeto que executa o Script do SAP GUI, desde o Login do usuário
        Pode estar já logado ou não.
        Atributos:
            - users (str): Usuário do Login do SAP
            - password (str): Senha do Login do SAP
            - acess_name (str): Nome da conexão do acesso ao SAP
    """
    def __init__(self, users, password, acess_name):
        self.users = users
        self.password = password
        self.acess_name = acess_name
        # Abrir o SAP
        try:
            self.SapGuiAuto = win32.GetObject('SAPGUI')
            self.aplicativo = self.SapGuiAuto.GetScriptingEngine
            self.session = self.aplicativo.Children(0).Children(0)
        
        except:
            self.path = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
            subprocess.Popen(self.path)
            time.sleep(7) # PC que uso é muito lento, dependendo esse valor de espera pode ser menor

            self.SapGuiAuto = win32.GetObject('SAPGUI')
            aplicativo = self.SapGuiAuto.GetScriptingEngine
            self.connection = aplicativo.OpenConnection(self.acess_name, True)
            self.session = self.connection.Children(0)
            time.sleep(3)

            self.session.findById('wnd[0]').maximize
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.users
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
            self.session.findById("wnd[0]").sendVKey(0)

    def conexaoSap(self, transacao):
        '''
            Faz um acesso a transação fornecida pelo usuário
            Exibe uma mensagem de erro caso o Usuário não tenha permissão da transação
            Atributos:
                - transacao (str): Nome da transação do SAP 
        '''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/o"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = transacao
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
        except:
            print(f'Usuário não possui acesso para a transação {transacao}')

    def sapGetCabecalho(self):
        """
            Irá fornecer as ordens de produção na aba de Cabeçalho de ordens a partir da transação 'COOIS'
            Padrões de filtros já pré definidos para uso específico
        """
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_DISPO_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "z04"
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "ente"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "ence"
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")

    def sapGetComponentes(self, salvar, txt_name):
        """
            Irá fornecer as ordens de produção na aba Componentes a partir da transação 'COOIS'
            É necessário que o número das Ordens de produção estejam em um arquivo .txt e passar como parâmetro
            Atributos:
                - salvar (str): Local aonde está armazenado o arquivo txt com o número das Ordens de produção
                - txt_name (str): Nome do arquivo txt com as Ordens de produção
        """
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
        self.session.findById("wnd[2]/usr/ctxtDY_PATH").text = salvar
        self.session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = txt_name
        self.session.findById("wnd[2]").sendVKey(0)
        self.session.findById("wnd[1]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")


    def gerarPlanilha(self, salvar, nome):
        """
            Exporta uma planilha da transação 'COOIS'
            É necessário já estar dentro da transação para fornecer essa planilha
            Atributos:
                - salvar (str): Local onde será salva a planilha
                - nome (str): Nome da planilha para salvar com a Extensão do arquivo
        """
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = salvar
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome
        self.session.findById("wnd[1]").sendVKey(0)


    def GetCabecalhoSemPlanilha(self, destino):
        """
            Extrai as informações na transação "COOIS" sem a necessidade de gerar uma planilha no SAP
            As informações são armazenadas e filtradas por uma DateFrame através do Pandas e salva o arquivo no lugar específicado pelo Parâmetro
            Atributos:
                - destino (str): Local onde será salva a planilha já filtrada
        """
        colunas = ['Ordem', 'Material', 'Ordem do Cliente', 'Item ord.cliente', 'Texto breve material', 'Status do sistema', 'Data-base iníc.', 'Data conclusão (prog.)', 'Quantidade da ordem (GMEIN)', 'Quantidade boa confirmada (GMEIN)', 'Qtd.fornecida (GMEIN)', 'Unidade de medida (=GMEIN)', 'Depósito', 'Planejador MRP', 'Versão de produção', 'Data de entrada', 'Criado por', 'Change date', 'Último modificador', 'Dads.explosão lis.tarefas/LisTéc.', 'Data fim real']

        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_DISPO_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "z04"
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "ente"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "ence"
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn("IGMNG")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&MB_FILTER")
        self.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = ""
        self.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select()
        self.session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,0]").text = "0"
        self.session.findById("wnd[2]").sendVKey(8)
        self.session.findById("wnd[1]").sendVKey(0)

        myGrid = self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell")
        allRows = int(myGrid.RowCount) # Número de SAP Linhas
        allCols = int(myGrid.ColumnCount) # Número de SAP Colunas
        columns = myGrid.ColumnOrder #SAP column names in order in SAP window
        linha = []
        total_linhas = []

        for j in range(allRows):
            myGrid.firstVisibleRow = j
            for i in range(allCols):
                linha.append(myGrid.GetCellValue(j, columns(i)))
                #Cells(j + 1, i + 1).Value = myGrid.GetCellValue(j, columns(i))
            total_linhas.append(linha)
            linha = []
        planilha = pd.DataFrame(total_linhas, columns=colunas)
        planilha.insert(11, 'Qtde Falta', planilha['Quantidade da ordem (GMEIN)'] - planilha['Qtd.fornecida (GMEIN)'])
        remove_line = planilha[planilha['Qtde Falta'] > 0].index
        planilha = planilha.drop(remove_line)
        planilha.loc[planilha['Versão de produção'] == '0', 'Versão de produção'] = 0
        with open(f'{destino}\ordens.txt', 'w+') as txt:
            txt.write(planilha['Ordem'].to_string(index=False))
        planilha.to_excel(f'{destino}/Cabecalho.xlsx', sheet_name='Cabeçalho', index=False)

    
    def GetComponentesSemPlanilha(self, local, nome_txt):
        """
            Extrai os componentes das Ordens de Produção fornecidas
            As informações são lidas e armazenadas em um DateFrame Pandas e salvas no lugar específicado pelo parâmetro
            Atributos:
                - local (str): Local aonde está o arquivo .txt com os números das ordens, e também aonde será salva a planilha com os componentes
                - nome_txt (str): Nome do arquivo .txt que deve ser usado para capturar o número das Ordens de Produção
        """
        linha = []
        total_linhas = []
        colunas = ['Ordem', 'Material', 'Txt.brv.material', 'Lista comp.item', 'Requirement date', 'Qtd.necessária (EINHEIT)', 'Qtd.retirada (EINHEIT)', 'Unid.medida básica (=EINHEIT)', 'Depósito', 'Centro de trabalho', 'Descrição do centro de trabalho', 'Refugo da operação %', 'Status do sistema', 'Texto']
        
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
        self.session.findById("wnd[2]/usr/ctxtDY_PATH").text = local
        self.session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = nome_txt
        self.session.findById("wnd[2]").sendVKey(0)
        self.session.findById("wnd[1]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(8)       
        
        myGrid = self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell")
        allRows = int(myGrid.RowCount) - 1 # Número de SAP Linhas
        allCols = int(myGrid.ColumnCount) - 1 # Número de SAP Colunas
        columns = myGrid.ColumnOrder #SAP nomes de colunas em ordem na janela SAP

        for j in range(allRows):
            myGrid.firstVisibleRow = j
            for i in range(allCols):
                linha.append(myGrid.GetCellValue(j, columns(i)))
                #Cells(j + 1, i + 1).Value = myGrid.GetCellValue(j, columns(i))
            total_linhas.append(linha)
            linha = []

        
        planilha = pd.DataFrame(total_linhas, columns=colunas)
        planilha.insert(7, 'Qtde Falta', planilha['Qtd.necessária (EINHEIT)'] - planilha['Qtd.retirada (EINHEIT)'])
        planilha['Requirement date'] = planilha['Requirement date'].dt.strftime('%d/%m/%Y')
        planilha.to_excel(f'{local}/Componentes.xlsx', sheet_name='Componentes', index=False)


if __name__ == '__main__':
    test = SapGui('user', 'password', 'acess')
    test.conexaoSap('COOIS')