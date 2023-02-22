'''
Autora: Valesca Gomes Soares

Scripts de transações do SAP
'''

import win32com.client
from datetime import datetime, date

#Retorna as datas de confirmação de um local de instalação do TB01, TB02, TB05    
def notas(local):
    
    datas = {}
    hoje = date.today().strftime("%d.%m.%Y") #data de hoje
    
    #estabelecendo conexão com o SAP
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
      return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
      SapGuiAuto = None
      return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
      application = None
      SapGuiAuto = None
      return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
      connection = None
      application = None
      SapGuiAuto = None
      return

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw47"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkDY_ABG").selected = True
    session.findById("wnd[0]/usr/btn%_AUART_O_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN
    _HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]")
    .text = "TB01"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN
    _HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,1]")
    .text = "TB02"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN
    _HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,2]")
    .text = "TB05"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtSTRNO_O-LOW").text = local
    session.findById("wnd[0]/usr/ctxtERSDA_C-LOW").text = "01.01.2015"
    session.findById("wnd[0]/usr/ctxtERSDA_C-HIGH").text = hoje
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    status_bar = session.findById("wnd[0]/sbar").text

    if status_bar != "ID alternativo: nem todos os locais de instalação já foram convertidos":
        
        title = session.ActiveWindow.text
        
        if "Confirmação para ordem" in title:
            dia = session.findById("wnd[0]/usr/ctxtAFRUD-ISDD")
		.text.split('.')[0]
            mes = session.findById("wnd[0]/usr/ctxtAFRUD-ISDD")
		.text.split('.')[1]
            ano = session.findById("wnd[0]/usr/ctxtAFRUD-ISDD")
		.text.split('.')[2]
            data_sap = ano + "-" + mes + "-" + dia + "T15:00:00Z"
            
            session.findById("wnd[0]/usr/txtAFVGD-AUFNRD").setFocus()
            session.findById("wnd[0]").sendVKey(2)
            centro = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:
		3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/
		ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/
		ctxtCAUFVD-VAPLZ").text
            descricao = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:
		3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/
		subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").text
            
            
            if "MEC" in centro:
                if ( (("TROC" in descricao.upper()) and ("MOT" in descricao.upper()) and "92" in local) or (("TROC" in descricao.upper()) and ("BOMB" in descricao.upper())) and "93" in local ): 
                    datas[1] = data_sap
                session.findById("wnd[0]/tbar[0]/btn[3]").press() 
            elif ("ELE" in centro or "ANV" in centro):
                if ( ("LIMP CARCAÇA E INSP CX LIGAÇÃO" in descricao.upper()) or ("INSP. SENSITIVA MOTOR ELÉTRICO" in descricao.upper()) or
                    ("MONIT. PREDITIVO VIBRAÇÃO MOTOR ELÉTRICO" in descricao.upper()) or ("ENSAIO" in descricao.upper()) or 
                    ("MONIT. PREDITIVO TERMO. MOTOR ELÉTRICO" in descricao.upper()) or ("INSP. SENS. MOTOR ELÉTRICO S/ TERMOG." in descricao.upper()) or
                    ("PREV. SISTEMÁTICA MOTOR ELÉTRICO" in descricao.upper()) or ("PREV. SISTEM. MOTOR ELÉTRICO" in descricao.upper()) or
                    ("INSPEÇÃO MOT" in descricao.upper() and "92" in local) or ("INSPEÇÃO BOMB" in descricao.upper() and "93" in local) or 
                    ("Realizar ensaio para verificar") in descricao.upper()):
                     datas[1] = data_sap
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press() 
            session.findById("wnd[0]/tbar[0]/btn[3]").press()             
            
            return datas
    
        table = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        tam_tabela = table.RowCount
        linha = 0
      
        while linha < tam_tabela:
          if linha % 15 == 0:
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
		.firstVisibleRow = linha
          try: 
            if datetime.strptime(session.findById("wnd[0]/usr/
		cntlGRID1/shellcont/shell")
		.getCellValue(linha,"ISDD"), "%d.%m.%Y")
		 < datetime.strptime(hoje, "%d.%m.%Y"):
                dia = session.findById("wnd[0]/usr/cntlGRID1/shellcont
		    /shell").getCellValue(linha,"ISDD").split('.')[0]
                mes = session.findById("wnd[0]/usr/cntlGRID1/shellcont
		    /shell").getCellValue(linha,"ISDD").split('.')[1]
                ano = session.findById("wnd[0]/usr/cntlGRID1/shellcont
                /shell").getCellValue(linha,"ISDD").split('.')[2]
                ordem = session.findById("wnd[0]/usr/cntlGRID1/shellcont
		    /shell").getCellValue(linha,"AUFNR")
                data_sap = ano + "-" + mes + "-" + dia + "T15:00:00Z"
                
                centro = session.findById("wnd[0]/usr/cntlGRID1/shellcont
		    /shell").getCellValue(linha,"IARBPL")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
		    .currentCellRow = linha
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
		    .doubleClickCurrentCell()
                if "MEC" in centro:
                    session.findById("wnd[0]/usr/txtAFVGD-AUFNRD")
			  .setFocus()
                    session.findById("wnd[0]").sendVKey(2)
                    descricao = session.findById("wnd[0]/usr/subSUB_ALL:
			  SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB
			  _KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/
			  cntlLTEXT/shell").text
                    if ( (("TROC" in descricao.upper()) and ("MOT" in descricao.upper()) and "92" in local) or (("TROC" in descricao.upper()) and ("BOMB" in descricao.upper())) and "93" in local ): 
                        datas[ordem] = data_sap
                    session.findById("wnd[0]/tbar[0]/btn[3]")
			  .press() 
                    session.findById("wnd[0]/tbar[0]/btn[3]")
			  .press() 
                elif ("ELE" in centro or "ANV" in centro):
                    descricao = session.findById("wnd[0]/usr/txtCAUFVD-
			  KTEXT").text
                    if ( ("LIMP CARCAÇA E INSP CX LIGAÇÃO" in descricao.upper()) or ("INSP. SENSITIVA MOTOR ELÉTRICO" in descricao.upper()) or
                        ("MONIT. PREDITIVO VIBRAÇÃO MOTOR ELÉTRICO" in descricao.upper()) or ("ENSAIO" in descricao.upper()) or 
                        ("MONIT. PREDITIVO TERMO. MOTOR ELÉTRICO" in descricao.upper()) or ("INSP. SENS. MOTOR ELÉTRICO S/ TERMOG." in descricao.upper()) or
                        ("PREV. SISTEMÁTICA MOTOR ELÉTRICO" in descricao.upper()) or ("PREV. SISTEM. MOTOR ELÉTRICO" in descricao.upper()) or
                        ("INSPEÇÃO MOT" in descricao.upper() and "92" in local) or ("INSPEÇÃO BOMB" in descricao.upper() and "93" in local) ):
                         datas[ordem] = data_sap
                    session.findById("wnd[0]/tbar[0]/btn[3]")
			  .press()
                
            linha = linha+1
          except:
            datas = False
      
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
    else:
        datas = False
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
    
    return datas