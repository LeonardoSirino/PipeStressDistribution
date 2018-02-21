VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private data(), titles() As String

Public Sub ReadReport(ReportPath As String)
    Dim temp() As String
    Dim i, n As Integer
    ReDim titles(-1)
    
    Open ReportPath For Input As #1
    
    i = 0
    n = 1
    Do Until EOF(1)
        Line Input #1, Text
        MsgBox (Text)
        If Text = "&" Then
            n = n + 1
            i = 0
            Call InsertElement(data, temp)
            ReDim temp(-1)
        Else
            If i = 0 Then
                Call InsertElement(titles, Text)
            Else
                Call InsertElement(temp, Text)
            End If
            i = i + 1
        End If
    MsgBox (Str(data.Lenght))
End Sub


Private Sub InsertElement(Colection() As Variant, element As Variant)

        ReDim Preserve Colection(1 To UBound(Colection) + 1)
        Colection(UBound(Colection)) = element
        
End Sub

