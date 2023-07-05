Sub Solicitar()
'
' Solicitar abertura de serviços a contratada

' Desativar atualização de tela durante as macros
Application.ScreenUpdating = False
Application.DisplayAlerts = False


    ' Desbloqueando Abertura de OM
    ActiveSheet.Unprotect
    Sheets("Serviços").Select
    ' Desbloqueando Relatório
    ActiveSheet.Unprotect
    
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Range("A4").Select
    Sheets("Abrir Serviço").Select 'Voltar para aba de serviços
    Range("Cadastrar[[Num]:[Prioridade]]").Select
    Selection.Copy
    Sheets("Serviços").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H4").Select
    Sheets("Abrir Serviço").Select
    Range("Cadastrar[[Data de emissão]:[Prazo]]").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Serviços").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O4").Select
    Sheets("Abrir Serviço").Select
    Range("Cadastrar[Empresa]").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Serviços").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A4").Select
    
    ' Bloqueando a planilha de reatório
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
    ActiveSheet.EnableSelection = xlNoRestrictions
    
    ' Limpando espaços em branco
    Sheets("Abrir Serviço").Select
    Range("Cadastrar[[Solicitante]:[Prazo]]").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Cadastrar[Empresa]").Select
    Selection.ClearContents
    
    Range("Cadastrar[Data de emissão]").Select
    ActiveCell.FormulaR1C1 = "=IF([@[Descrição do serviço]]<>"""",TODAY(),"""")"
    Range("Cadastrar[Solicitante]").Select
    
    ' Bloqueando novamente a planiliha de abertura
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveSheet.EnableSelection = xlUnlockedCells

' Ativar atualização de tela durante as macros
Application.ScreenUpdating = False
Application.DisplayAlerts = False

End Sub

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

Sub PDFexport()

Dim intervalo As Range
Set intervalo = Sheets("Imprimir").Range("A3:H40")

NomeArquivo = Application.GetSaveAsFilename(InitialFileName:=Range("D42").Value, _
FileFilter:="PDF, *.pdf", _
Title:="Salve as PDF")

If NomeArquivo <> False Then

intervalo.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Range("D42").Value
End If

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

End Sub
Sub Imprimir()

'Dim intervalo As Range
'Set intervalo = Sheets("Imprimir").Range("A3:H40")

ActiveSheet.PageSetup.PrintArea = "$A$3:$h$40"
ActiveWindow.SelectedSheets.PrintPreview

End Sub

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

Sub Verificar_vazias()
'
' Verificação de campos vazios antes de solicitar
'

Dim flag As Integer
flag = 0

If Range("B6").Value = "" Then
    MsgBox "O campo SOLICITANTE não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If Range("C6").Value = "" Then
    MsgBox "O campo DESCRIÇÃO não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If Range("D6").Value = "" Then
    MsgBox "O campo PREDIO não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If Range("E6").Value = "" Then
    MsgBox "O campo COMPLEMENTO não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If Range("F6").Value = "" Then
    MsgBox "O campo PRIORIDADE não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If Range("G6").Value = "" Then
    MsgBox "O campo DATA DE EMISSÃO não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If Range("H6").Value = "" Then
    MsgBox "O campo PRAZO não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If Range("J6").Value = "" Then
    MsgBox "O campo EMPRESA não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If flag = 0 Then
    Call Solicitar
    MsgBox "Dados salvos com sucesso", 64, "Gravação efetuada"
End If


End Sub