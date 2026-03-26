If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

' -------------------------------
' PARAMETERS FROM PYTHON
' -------------------------------
If WScript.Arguments.Count < 3 Then
    MsgBox "Missing arguments"
    WScript.Quit
End If

contractType = WScript.Arguments(0)
salesOrg     = WScript.Arguments(1)
customer     = WScript.Arguments(2)

session.findById("wnd[0]").resizeWorkingPane 69,22,false

' Start VA41
session.findById("wnd[0]/tbar[0]/okcd").text = "VA41"
session.findById("wnd[0]").sendVKey 0

' Header
session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = contractType
session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = salesOrg
session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = "TR"
session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = "CR"

session.findById("wnd[0]/usr/ctxtVBAK-SPART").setFocus
session.findById("wnd[0]/usr/ctxtVBAK-SPART").caretPosition = 2
session.findById("wnd[0]/tbar[0]/btn[0]").press

' Customer
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = customer
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = customer

' Dates (leave static for now)
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBAK-GUEBG").text = "03/25/2026"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBAK-GUEEN").text = "03/31/2026"

' Item
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/txtVBAP-POSNR[0,0]").text = "10"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/ctxtRV45A-MABNR[1,0]").text = "CRUACCE01"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/txtVBAP-ZMENG[3,0]").text = "1000.000"

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/ctxtRV45A-MABNR[1,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/ctxtRV45A-MABNR[1,0]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/ctxtRV45A-MABNR[1,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/ctxtRV45A-MABNR[1,0]").caretPosition = 6
session.findById("wnd[0]").sendVKey 2


' Storage location
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03/ssubSUBSCREEN_BODY:SAPMV45A:4452/ctxtVBAP-LGORT").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03/ssubSUBSCREEN_BODY:SAPMV45A:4452/ctxtVBAP-LGORT").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/tbar[0]/btn[0]").press

' Incoterms
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\04").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4453/ctxtVBKD-INCO1").text = "DDP"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4453/ctxtVBKD-INCO1").caretPosition = 3
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

' TSW Data
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN").getAbsoluteRow(0).selected = true
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/txtVBAP-POSNR[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/txtVBAP-POSNR[0,0]").caretPosition = 4
session.findById("wnd[0]").resizeWorkingPane 144,23,false
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/subSUBSCREEN_BUTTONS:SAPMV45A:4052/btnOI_OIAC").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN").getAbsoluteRow(0).selected = true
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/txtVBAP-POSNR[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN/txtVBAP-POSNR[0,0]").caretPosition = 4
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/subSUBSCREEN_BUTTONS:SAPMV45A:4052/btnTSW_DATA_PB").press
session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabpOIJDET/ssubSUBSCREEN_BODY:SAPLOIJ_EL_C:0111/chkOIJ_EL_DOC_MOT-TSWRELV").selected = true
session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabpOIJDET/ssubSUBSCREEN_BODY:SAPLOIJ_EL_C:0111/ctxtOIJ_EL_DOC_MOT-PLANLOC").text = "GEN_HUS"
session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabpOIJDET/ssubSUBSCREEN_BODY:SAPLOIJ_EL_C:0111/ctxtOIJ_EL_DOC_MOT-TSYST").text = "ATA_MARINE"
session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabpOIJDET/ssubSUBSCREEN_BODY:SAPLOIJ_EL_C:0111/chkOIJ_EL_DOC_MOT-TSWRELV").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

' Save
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press