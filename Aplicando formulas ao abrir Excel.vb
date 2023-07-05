Private Sub Workbook_Open()

    Application.ScreenUpdating = False
    
    For i = 2 To Sheets.Count
        Sheets(i).Select
		If ActiveSheet.Name = "EAP" Then Exit For
		
        ActiveSheet.Unprotect
        Call ApplyFormula
        Call ApplyFormula2
        
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, _
        AllowFiltering:=True
    Next i
    
    Sheets("EAP").Select
    ActiveSheet.Protect
    
    Sheets("PREVIEW").Select
    ActiveSheet.Protect
    Range("C2").Select
    
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Sub ApplyFormula()
' Exec formula nas celulas em cada sheet
'
Dim ultCellProcess As Range
    
    Range("B6").Select
    Do While ActiveCell <> ""
        Set ultCellProcess = ActiveCell
        'ActiveCell.Offset(0, 1).Value
        
        ActiveCell.Offset(0, 1).FormulaR1C1 = _
        "=IF([@SS]<>"""",CONCAT([@SS],"" - "",[@Serviços]),"""")"

        ultCellProcess.Select
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Sub ApplyFormula2()
    
    Range("N6").Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(VLOOKUP([@[Item EAP]],TableEAP,2,FALSE),"""")"
    Range("N7").Select
End Sub

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------