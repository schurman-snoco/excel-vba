Sub errorHandling()

On Error GoTo errHandler
    'Code goes here

exit_errorHandling:
    'Error handling code goes here
    Exit Sub
    
errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    Resume exit_errorHandling

End Sub
