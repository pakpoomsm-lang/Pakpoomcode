If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUISERVER")
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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP0022"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").resizeWorkingPane 194,25,false
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "1001"
session.findById("wnd[0]/usr/radR_PROC_3").setFocus
session.findById("wnd[0]/usr/radR_PROC_3").select
session.findById("wnd[0]/usr/ctxtS_LINE3-LOW").text = "542"
session.findById("wnd[0]/usr/ctxtS_LINE3-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_LINE3-LOW").caretPosition = 3
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 33
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").setCurrentCell 10,"TEXT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").firstVisibleRow = 3
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "10"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 5,"PSMNG"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "5"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[20]").press
On Error Resume Next
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "J:\\7.541_HEI\\Database follow\\ZPP0022"
session.findById("wnd[2]/usr/ctxtDY_PATH").text = "J:\\7.541_HEI\\Database follow\\ZPP0022"
On Error GoTo 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
