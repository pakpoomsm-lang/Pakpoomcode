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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP0059"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").resizeWorkingPane 153,29,false
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_SHOP-LOW").text = "542"
session.findById("wnd[0]/usr/ctxtS_ORDTY-LOW").text = "zp40"
session.findById("wnd[0]/usr/ctxtS_WKDT-LOW").text = "11.06.2026"
session.findById("wnd[0]/usr/ctxtS_WKDT-HIGH").text = "18.06.2026"
session.findById("wnd[0]/usr/ctxtS_WKDT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtS_WKDT-HIGH").caretPosition = 10
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 2,"ZTIMESTAMP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-DESTINATION").setFocus
session.findById("wnd[1]/tbar[0]/btn[20]").press
On Error Resume Next
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "J:\\7.541_HEI\\Database follow\\ZPP0059"
session.findById("wnd[2]/usr/ctxtDY_PATH").text = "J:\\7.541_HEI\\Database follow\\ZPP0059"
On Error GoTo 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
