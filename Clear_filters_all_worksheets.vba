Sub Auto_Open()
    Dim xWs As Worksheet
    For Each Wks In ThisWorkbook.Worksheets
    On Error Resume Next
    If Wks.AutoFilterMode Then
        Wks.AutoFilterMode = False
    End If
    Next Wks
End Sub