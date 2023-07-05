Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.ScreenUpdating = False
For i = 1 To Sheets.Count
    Sheets(i).Select
    
    If ActiveSheet.Protect = False Then
        ActiveSheet.Protect
    'ActiveSheet.EnableSelection = xlNoRestrictions
    End If
Next i
Sheets("PREVIEW").Select
Range("C2").Select
End Sub