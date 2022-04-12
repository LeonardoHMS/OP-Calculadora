# Classe para acesso ao SAP, em construção, ainda não testado !!!
import win32com.client as win32
import subprocess
import time

class SapGui(object): # Classe para abrir o sistema SAP
    def __init__(self):
        # Abrir o SAP
        self.path = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
        subprocess.Popen(self.path)

        self.SapGuiAuto = win32.GetObject('SAPGUI')
        aplicativo = self.SapGuiAuto.GetScriptingEngine     
        self.connection = aplicativo.OpenConnection('Nome de Acesso', True) # Informe o nome de acesso ao SAP
        time.sleep(5)
        self.session = self.connection.Children(0)
        self.session.findById('wnd[0]').maximize


    def conexaoSap(self, transacao): # Loga no Sap e entra na transação fornecida
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "usuario" # Informe seu usuário de login
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "senha"   # Informe sua senha de login
            self.session.findById("wnd[0]").sendVKey(0)
        except:
            pass
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/o"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[0]/okcd").text = transacao # Transação que deseja abrir
        self.session.findById("wnd[0]").sendVKey(0)


    def SapCooisXlsx(self): 
        # Function para gerar planilha, logo faço como classe filha !!!
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-LOW").setFocus
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-LOW").caretPosition = 0
        self.session.findById("wnd[0]").sendVKey(2)
        self.session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton("EXCL")
        self.session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = True
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-LOW").text = "z04"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "ente"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "ence"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").setFocus
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellColumn = "MATXT"
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\\Users\racao01\Desktop" # Local aonde será salvo o arquivo
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 4
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press


if __name__ == '__main__':
    Sap_test = SapGui()
    Sap_test.sapLogin()
    Sap_test.conexaoSap('COOIS')
    Sap_test.SapCooisXlsx()
