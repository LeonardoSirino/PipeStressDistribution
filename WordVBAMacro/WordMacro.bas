Attribute VB_Name = "M�dulo1"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub macroSirino()

Dim report As New report

    report.ReadReport ("Tabelas.txt")
    
End Sub
