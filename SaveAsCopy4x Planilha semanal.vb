Sub ExportarWorkbook()

    Dim LocArqv As Range
    Dim NameArqv As Range
    Dim dataAnterior As Date
    
    Set LocArqv = Sheets("Mês de Medição - Período").Range("M2")
    Set NameArqv = Sheets("Mês de Medição - Período").Range("J10")
    
    dataAnterior = Range("E28").Value
    
    If NameArqv <> False Then
        ActiveWorkbook.SaveCopyAs LocArqv & NameArqv
    End If
    
Application.ScreenUpdating = False
Sheets("Mês de Medição - Período").Visible = True
'---------------------------------------------------
    For i = 1 To 3
        Sheets("Mês de Medição - Período").Select
        Range("M6").Select
        Selection.Copy
        Sheets("Relatório de Atividades").Select
        Range("E28").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("A14:J19").Select
        
        If NameArqv <> False Then
            ActiveWorkbook.SaveCopyAs LocArqv & NameArqv
        End If
    Next i
'---------------------------------------------------
Sheets("Mês de Medição - Período").Visible = False

Range("E28").Value = dataAnterior
Application.ScreenUpdating = True
End Sub