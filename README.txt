'Error Handling for a button which throws error

'Error handler is the error message displayed by VBA

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
