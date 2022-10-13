# -------------------------
# Classe para acesso ao SAP
# Abertura e Login automático do FrontEnd SAP
# Integração com SAP GUI Scripting API
# Leonardo Mantovani github.com/LeonardoHMS
# -------------------------
import win32com.client as win32
from time import sleep
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
        Metodos:
            conexaoSAP: Entra na janela da transação fornecida para acessar
            - sapGetCabecalho: A partir da transação 'COOIS' irá pegar o Cabeçalho de ordens
                 utilizando alguns filtros de pesquisa
            - sapGetComponentes: A partir da transação 'COOIS irá pegar os Componentes de ordens
                 utilizando alguns filtros de pesquisa
            - gerarPlanilha: Exporta uma planilha da transação 'COOIS'
                atributos:
                    - salvar (str): Diretório para salvar a planilha
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
            time.sleep(3)

            self.SapGuiAuto = win32.GetObject('SAPGUI')
            aplicativo = self.SapGuiAuto.GetScriptingEngine
            while True:
                try:
                    self.connection = aplicativo.OpenConnection(self.acess_name, True)
                    break
                except:
                    sleep(2)
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


    def gerarPlanilha(self, salvar, nome):
        """
            Exporta uma planilha da transação 'COOIS'
            É necessário já estar dentro da transação para fornecer essa planilha
            Atributos:
                - salvar (str): Local a onde será salva a planilha
                - nome (str): Nome da planilha para salvar com a Extensão do arquivo
        """
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = salvar
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome
        self.session.findById("wnd[1]").sendVKey(0)
