Sub GetEachWorksheetSize()
    Dim strTargetSheetName As String
    Dim strTempWorkbook As String
    Dim objTargetWorksheet As Worksheet
    Dim objWorksheet As Worksheet
    Dim objRange As Range
    Dim i As Long
    Dim nLastEmptyRow As Integer

    strTargetSheetName = "Sheet Sizes"
    strTempWorkbook = ThisWorkbook.Path & "\Temp Workbook.xls"

    With ActiveWorkbook.Worksheets.Add(Before:=Application.Worksheets(1))
         .Name = strTargetSheetName
         .Cells(1, 1) = "Sheet"
         .Cells(1, 1).Font.Size = 14
         .Cells(1, 1).Font.Bold = True
         .Cells(1, 2) = "Size"
         .Cells(1, 2).Font.Size = 14
         .Cells(1, 2).Font.Bold = True
    End With

    Set objTargetWorksheet = Application.Worksheets(strTargetSheetName)

    For Each objWorksheet In Application.ActiveWorkbook.Worksheets
        If objWorksheet.Name <> strTargetSheetName Then
           objWorksheet.Copy

           Application.ActiveWorkbook.SaveAs strTempWorkbook
           Application.ActiveWorkbook.Close SaveChanges:=False

           nLastEmptyRow = objTargetWorksheet.Range("A" & objTargetWorksheet.Rows.Count).End(xlUp).Row + 1

           With objTargetWorksheet
                .Cells(nLastEmptyRow, 1) = objWorksheet.Name
                .Cells(nLastEmptyRow, 2) = FileLen(strTempWorkbook)
           End With

           Kill strTempWorkbook
         End If
    Next
End Sub