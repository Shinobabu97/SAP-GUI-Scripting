Option Explicit
Public SapGuiAuto, WScript, msgcol
Public objGui As GuiApplication
Public objConn As GuiConnection
Public objSBar As GuiStatusbar
Public session As GuiSession

Sub Nettowert_from_SAP()

    ' Initialize SAP connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set objGui = SapGuiAuto.GetScriptingEngine
    Set objConn = objGui.Children(0)
    Set session = objConn.Children(0)

    Dim nettowert As String
    Dim ws As Worksheet
    Dim Lieferungsnummer As String
    Dim lastRow As Long
    Dim i As Long
    Dim statusText As String

    Set ws = ThisWorkbook.Sheets("Tabelle1")
    Set objSBar = session.FindById("wnd[0]/sbar")
    


    ' Find last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in column A
    For i = 1 To lastRow
        Lieferungsnummer = ws.Cells(i, 1).Value  ' Get value from column A

        ' Skip if the cell is empty
        If Lieferungsnummer <> "" Then

            ' SAP actions
            session.FindById("wnd[0]").Maximize
            session.FindById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").DoubleClickNode "F00002"
            session.FindById("wnd[0]/usr/radP_INEX").Select
            session.FindById("wnd[0]/usr/chkP_KEINWA").Selected = False
            session.FindById("wnd[0]/usr/ctxtS_VBELN-LOW").Text = Lieferungsnummer
            session.FindById("wnd[0]/usr/chkP_KEINWA").SetFocus
            session.FindById("wnd[0]/tbar[1]/btn[8]").Press
            If objSBar.Text = "Keine Lieferungen zu den eingegebenen Auswahlkriterien gefunden!" Then
                ws.Cells(i, 2).Value = "-"
                session.FindById("wnd[0]/tbar[0]/btn[15]").Press
                GoTo handler
            End If
            session.FindById("wnd[0]/usr/lbl[19,1]").SetFocus
            session.FindById("wnd[0]/usr/lbl[19,1]").CaretPosition = 1
            session.FindById("wnd[0]").SendVKey 2
            session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetCurrentCell 2, "VBELN"
            session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell
            session.FindById("wnd[0]/tbar[1]/btn[7]").Press
            session.FindById("wnd[0]/usr/shell/shellcont[1]/shell[1]").SelectItem "          1", "&Hierarchy"
            session.FindById("wnd[0]/usr/shell/shellcont[1]/shell[1]").EnsureVisibleHorizontalItem "          1", "&Hierarchy"
            session.FindById("wnd[0]/tbar[1]/btn[8]").Press

            'Special error case
            statusText = objSBar.Text
            If Left(statusText, 25) = "Keine Anzeigeberechtigung" Then
                ws.Cells(i, 2).Value = "-"
                session.FindById("wnd[0]/tbar[0]/btn[12]").Press
                session.FindById("wnd[0]/tbar[0]/btn[15]").Press
                session.FindById("wnd[0]/tbar[0]/btn[15]").Press
                session.FindById("wnd[0]/tbar[0]/btn[15]").Press
                session.FindById("wnd[0]/tbar[0]/btn[15]").Press
                session.FindById("wnd[0]/tbar[0]/btn[15]").Press
                GoTo handler
            End If

            ' Handle possible error
            On Error Resume Next ' This deactivates the error handler and executes the GoTo which otherewise would have been an error
            session.FindById("wnd[1]/tbar[0]/btn[0]").Press
            If Err.Number <> 0 Then
                Err.Clear
                GoTo Myerror
            End If
            On Error GoTo 0 ' If there was no error, the handler is reactivated for error handling the further code
            
            
            
            ' Retrieve value from SAP
            session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtRV45A-ZZORDVAL").SetFocus
            nettowert = session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtRV45A-ZZORDVAL").Text
            ws.Cells(i, 2).Value = nettowert  ' Store result in column B
           
           
Myerror:
    ' Handle errors by still attempting to get the value from SAP
    session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtRV45A-ZZORDVAL").SetFocus
    nettowert = session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtRV45A-ZZORDVAL").Text
    ws.Cells(i, 2).Value = nettowert
        
        End If
        session.FindById("wnd[0]/tbar[0]/btn[15]").Press
        session.FindById("wnd[0]/tbar[0]/btn[15]").Press
        session.FindById("wnd[0]/tbar[0]/btn[15]").Press
        session.FindById("wnd[0]/tbar[0]/btn[15]").Press
        session.FindById("wnd[0]/tbar[0]/btn[15]").Press
        session.FindById("wnd[0]/tbar[0]/btn[15]").Press
handler:
    Next i

    
End Sub


