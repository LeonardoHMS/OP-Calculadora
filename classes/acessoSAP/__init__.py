# -------------------------
# Classe para acesso ao SAP
# Abertura e Login automático do FrontEnd SAP
# Leonardo Mantovani github.com/LeonardoHMS
# -------------------------
import win32com.client as win32
import subprocess
import time

class SapGui(object): # Classe para abrir o sistema SAP
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
            self.connection = aplicativo.OpenConnection(self.acess_name, True) # Informe o nome de acesso ao SAP
            self.session = self.connection.Children(0)
            time.sleep(3)

            self.session.findById('wnd[0]').maximize
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.users # Informe seu usuário de login
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password # Informe sua senha de login
            self.session.findById("wnd[0]").sendVKey(0)


    def conexaoSap(self, transacao): # Entra na transação fornecida
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/o"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = transacao # Transação que deseja abrir
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
        except:
            print(f'Usuário não possui acesso para a transação {transacao}')


    def sapCooisXlsx(self, salvar):
        # Function para gerar planilha
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_DISPO_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "z04"
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "ente"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "ence"
        self.session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)

        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = salvar #Local aonde irá salvar a planilha
        self.session.findById("wnd[1]").sendVKey(0)


if __name__ == '__main__':
    Sap_test = SapGui('usuario', 'senha', 'acesso')
    Sap_test.conexaoSap('COOIS')
    Sap_test.sapCooisXlsx('C:/Users/Leonardo Mantovani/Desktop')
