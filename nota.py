'''
Autora: Valesca Gomes Soares

Scripts de transações do SAP
'''

import win32com.client

def aberturaNota(local_instalacao, area, descricao):
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

    session.findById("wnd[0]").resizeWorkingPane(98,18,False)
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw21"
    session.findById("wnd[0]").sendVKey(0)

    #tipo de nota (T1 - nota do mantenedor)
    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = "T1"
    session.findById("wnd[0]").sendVKey(0)
    
    local_instalacao = str(local_instalacao)
    area = str(area)
    
    #título para nota , OBS: máximo de 40 caracteres pegar Portuguese Asset Name
    if(len(descricao) > 40):
      titulo = descricao[0:40]
      session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/
	txtVIQMEL-QMTXT").text = titulo
      descricao = descricao[40:]
    else:
        session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/
	  txtVIQMEL-QMTXT").text = descricao
        
    #Codificação (B103 / 0 - fora de parada)
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_2:SAPLIQS0:7710/ctxtVIQMEL-QMGRP").text = "B103"
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_2:SAPLIQS0:7710/ctxtVIQMEL-QMCOD").text = "0"
    
    
    if area == "PM4" or "PM3" or "WY":
        if ".92." in local_instalacao:
            if "preditiva" in titulo:
                areanome = "LUBA1"
                plnj = "004"
            else:
                areanome = "ELEA1"
                plnj = "001"
        else:
            if "preditiva" in titulo:
                areanome = "LUBA1"
                plnj = "004"
            else:
                areanome = "MECA1"
                plnj = "003"
            
    elif area == "TG1" or "TG2" or "RB3" or "PB3" or "LK" or "WT" or "EVAP" or "SOAP" or "BOP" or "CONC" or "UTIL" or "CT" or "ET" or "CAUST" or "LK2" or "PUHW" or "CR3" or "NCG":
        if ".92." in local_instalacao:
            if "preditiva" in titulo:
                areanome = "LUBA2"
                plnj = "004"
            else:
                areanome = "ELEA2"
                plnj = "001"
        else:
            if "preditiva" in titulo:
                areanome = "LUBA2"
                plnj = "004"
            else:
                areanome = "MECA2"
                plnj = "003"
    else:
        if ".92." in local_instalacao:
            if "preditiva" in titulo:
                areanome = "LUBA2"
                plnj = "004"
            else:
                areanome = "ELEA2"
                plnj = "001"
        else:
            if "preditiva" in titulo:
                areanome = "LUBA2"
                plnj = "004"
            else:
                areanome = "MECA2"
                plnj = "003"
        
    #Grupo de planejamento (003 - mecanica / B103) (001 - eletrica / B103) (004 - lubrificacao / B103)
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:
    7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtVIQMEL-INGRP").text = plnj
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:
    7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtVIQMEL-IWERK").text = "B103"
     
    #Centro de trabalho responsável (MECA1 / B103) - OBS: MECA é a equipe e 1 é a área
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_3:SAPLIQS0:7326/ctxtRIWO00-GEWRK").text = areanome
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_3:SAPLIQS0:7326/ctxtRIWO00-SWERK").text = "B103"
    #autor da nota
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_3:SAPLIQS0:7326/txtVIQMEL-QMNAM").text = "PISystem"
    #prioridade
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_5:SAPLIQS0:7330/cmbVIQMEL-PRIOK").key = "3"
     
    #determinar as datas novamente? não
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT").verticalScrollbar
    .position = 1
    if session.ActiveWindow.Name == "wnd[1]":
      session.findById("wnd[1]/usr/btnBUTTON_2").press()
      
    scroll = 2
    linha = 1
    #texto descritivo
    if(len(descricao) > 72):
        while(len(descricao) > 72):
            session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\
		TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:
		SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT/
		txtLTXTTAB2-TLINE[0,"+str(linha)+"]").text = descricao[0:72]
            session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\
		TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:
		SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT")
		.verticalScrollbar.position = scroll
            descricao = descricao[72:]
            linha = linha + 1
            scroll = scroll + 1
            if linha == 3:
              linha = 1
        session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
	  ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:
	  7212/subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT/
 	  txtLTXTTAB2-TLINE[0,"+str(linha)+"]").text = descricao
        session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
	  ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:
	  7212/subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT/
	  txtLTXTTAB2-TLINE[0,"+str(linha+1)+"]")
	  .text = "Nota criada automaticamente pelo PISystem."
    else:
        session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
	  ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
	  subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT/txtLTXTTAB2-TLINE
	  [0,"+str(linha)+"]").text = descricao
        session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
	  ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
	  subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT/txtLTXTTAB2-TLINE
	  [0,"+str(linha+1)+"]").text = "Nota criada automaticamente pelo PISystem."

    #local de instalação
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT")
    .verticalScrollbar.position = 0
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
    ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
    subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:0140/ctxtRIWO1-TPLNR")
    .text = local_instalacao
    #comparar dados de objetos? não
    session.findById("wnd[0]").sendVKey(0)
    if session.ActiveWindow.Name == "wnd[1]":
      session.findById("wnd[1]/tbar[0]/btn[12]").press() #nao
      #session.findById("wnd[1]/tbar[0]/btn[0]").press() #sim
       
    #informação objeto
    if session.ActiveWindow.Name == "wnd[1]":
      session.findById("wnd[1]/usr/btnRIHEA-STRUKTUR").press()
      session.findById("wnd[1]/tbar[0]/btn[0]").press()
    else:
      session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/
	ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/
	subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:0140/
	btnOBJECTINFORMATION").press()
      session.findById("wnd[1]/usr/btnRIHEA-STRUKTUR").press()
      session.findById("wnd[1]/tbar[0]/btn[0]").press()
     
    #definir status do usuário - A013 confiabilidade
    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/
    btnANWENDERSTATUS").press()
    session.findById("wnd[1]/usr/btnODOWN").press()
    session.findById("wnd[1]/usr/btnODOWN").press()
    session.findById("wnd[1]/usr/sub:SAPLBSVA:0201[1]/chkJ_STMAINT-
    ANWSO[0,0]").selected = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
     
    #caso ocorra uma mensagem de erro no statusbar
    if session.findById("wnd[0]/sbar").messagetype == "E":
      return "erro: " + session.findById("wnd[0]/sbar").text
     
    status_bar = session.findById("wnd[0]/sbar").text

    nota = status_bar.split(' ')[1]
    session.findById("wnd[0]/tbar[0]/btn[12]").press()

    return nota