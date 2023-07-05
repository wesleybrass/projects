Sub GerarRelatorioSS()

Dim Linha As Long
Dim DiaRDO As String
Dim UltCellProcess  As Range
Dim TabelaRelatorio As ListObject
Set TabelaRelatorio = ActiveSheet.ListObjects("TabelaRelatorio")

'___O processo só iniciará caso a AnswerYes = True
Dim AnswerYes As String
Dim AnswerNo As String

AnswerYes = MsgBox("Deseja mesmo gerar um relatório?", vbYesNo, "Verificação")
If AnswerYes = vbYes Then
    
    '___START
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    
'___Tratamento de erros
On Error Goto msgError
    
    '___Retirando filtros e limpando tabela
    Sheets("Relatório Mensal").Select
    TabelaRelatorio.HeaderRowRange.AutoFilter
    TabelaRelatorio.HeaderRowRange.AutoFilter
    Range("TabelaRelatorio").Select
    Selection.ClearContents

'___Evitando erros com celulas iniciais vazias
Range("A1").Value = "1"
Range("A2").Value = "2"
Range("A3").Value = "3"


    '___Loop iniciado em cada planilha__CopyPaste_________________________________
    For i = 2 To Sheets.Count
        Sheets(i).Select
        If ActiveSheet.Name = "EAP" Then Exit For
        DiaRDO = Range("K1").Value 'Recebe o dia do RDO
        
        '___Validação se há celula vazia antes do relatorio
        If Range("B6").Value <> "" Then
            Range("B6:O20").Select 'Copiando toda a tabela (Range limitada)
            Selection.Copy
            
            Sheets("Relatório Mensal").Select
            Range("A4").Select
            '___Identificando proxima Célula vazia
            Dim xCell As Range
            For Each xCell In ActiveSheet.Columns(1).Cells
                If Len(xCell) = 0 Then
                    xCell.Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                        
                    '___Aplicando o dia do RDO na mesmas linhas coladas
                    xCell.Select
                    Do While ActiveCell.Value <> ""
                        Set UltCellProcess = ActiveCell
                        ActiveCell.Offset(0, 3).Value = "RDO do dia " & DiaRDO
                        UltCellProcess.Select
                        ActiveCell.Offset(1, 0).Select
                    Loop
                    Set UltCellProcess = Nothing
                    
                    Exit For
                End If
            Next
        End If
    Next i
    '___Fim do loop_______________________________________________________________


    '___Retornando a aba relatório mensal
    Sheets("Relatório Mensal").Select
    Range("A2").Select
    
    '___Removendo as colunas inúteis
    Columns("E:J").Select
    Selection.Delete 'Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Delete 'Shift:=xlToLeft
    
    '___Excluindo linhas vazias
    With Planilha34
        For Linha = .Cells(.Rows.Count, "A").End(xlUp).Row To 5 Step -1
            If .Cells(Linha, "A") = "" Or .Cells(Linha, "D") = "Não executado" Then
                .Rows(Linha).Delete
            End If
        Next Linha
    End With
    
    '___Arrumando tamanho das colunas
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 15
    Columns("D:D").ColumnWidth = 15
    Columns("E:E").ColumnWidth = 12
    Columns("F:F").ColumnWidth = 45
    Columns("G:G").ColumnWidth = 35
    Columns("H:H").ColumnWidth = 25
    
    '___Renomeando titulos das colunas
    Range("B4").Value = "Etapa SS"
    Range("C4").Value = "Dia do RDO"
    Range("D4").Value = "Status"
    Range("E4").Value = "Item"
    Range("F4").Value = "Descrição do item"
    Range("G4").Value = "Cálculo do item EAP"
    Range("H4").Value = "Observações"
    
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    MsgBox "Relatório concluído."
Else
    MsgBox ("Relatório interrompido")
End If
Exit Sub
msgError: MsgBox "Houve um erro ao fazer o relatório."

End Sub