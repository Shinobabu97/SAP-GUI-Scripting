#######################################################
Error Handling for a button which throws error
Error handler is the error message displayed by VBA
#######################################################

Set objSBar = session.FindById("wnd[0]/sbar") 'Status bar

On Error Resume Next ' This deactivates the error handler
        session.FindById("wnd[0]/usr/lbl[19,1]").SetFocus 'This is the button throwing the error
        If Err.Number <> 0 Then
            Err.Clear
            GoTo Myerror
        End If
        On Error GoTo 0 'if there were no errors, then error handler is turned back on
        
Myerror:
    If objSBar.Text = "Keine Lieferungen zu den eingegebenen Auswahlkriterien gefunden!" Then
        MsgBox "Keine Lieferungen zu den eingegebenen Auswahlkriterien gefunden!", vbInformation, "Keine Daten"
        session.FindById("wnd[0]/tbar[0]/btn[3]").Press 'Back button
    End If

################################################
Accessing Information Window
################################################

    ' Try to access the modal dialog window
    Dim popupText As String
    Dim popupWindow As Object
    Set popupWindow = session.FindById("wnd[1]")
    
    If Not popupWindow Is Nothing Then
    ' Access the text element - the exact ID may vary but this is a common pattern
    Dim textElement As Object
    Set textElement = popupWindow.FindById("usr/txtMESSTXT1") 'Can also check the exact id from the recording Script
    
        If Not textElement Is Nothing Then
            popupText = textElement.Text ' Get the text
        End If
    End If
    MsgBox "Extracted text: " & popupText

###################################################
