Sub RelatorioPreventivas_cart1()
'
' Macro para gerar relatorio

Dim ws As Excel.Worksheet
Dim carteira1 As ListObject
Set carteira1 = ActiveSheet.ListObjects("Tabela1")

'------------------> O processo só iniciará caso a AnswerYes = True
Dim AnswerYes As String
Dim AnswerNo As String

AnswerYes = MsgBox("Você já confirmou todas as Ordens/Notas executadas dentro deste período? (Ordens confirmadas após o dia 15, deverá ter status LIB)", vbYesNo, "Verificação")
If AnswerYes = vbYes Then
'--v--v--v--v--v--v--v--v-- Executar codigo abaixo. Completo!
Application.ScreenUpdating = False
Application.CutCopyMode = False


' START
Sheets("Relatório").Select ' <---
'------------------> Fazer limpeza da planilha de relatório

    ' Primeiro limparemos o conteúdo
    Range("TabelaRelatorio").Select
    Selection.ClearContents


Sheets("01-Carteira").Select ' <---
'------------------> Fazer filtros e critérios necessário

' Limpar possíveis filtros deixados pelo usuário antes do processo
carteira1.HeaderRowRange.AutoFilter
carteira1.HeaderRowRange.AutoFilter

    ' Filtrando as preventivas
    Range("Tabela1").AutoFilter Field:=10, Criteria1:="SAP Pre*"
    
    ' Filtrando (Field 5 = Data de Liberação)
    Range("Tabela1").AutoFilter Field:=5, Criteria1:= _
    "<=" & Format(Planilha0.Range("DataFinal").Value, "mm/dd/yyyy hh:mm")


'------------------> Transferir dados da Carteira, para o Relatório

Range("Tabela1").Select ' Seleciona a tabela
On Error Goto msgError ' Caso não haja linhas para copiar, salta a linha
Selection.SpecialCells(xlCellTypeVisible).Copy ' Copiando...

    
Sheets("Relatório").Select ' <---
'------------------> Voltando à planilha de relatório
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


'------------------------------------------------------------------
'------------------> Alterando STATUS das ordens <-----------------

Dim w       As Worksheet
Dim UltCellProcess  As Range

Set w = Sheets("Relatório")
w.Range("A6").Select

Do While ActiveCell.Value <> ""
    Set UltCellProcess = ActiveCell
    If ActiveCell.Offset(0, 6).Value <> "" Then
            ActiveCell.Offset(0, 7).Value = "ENTE CONF IMPR CAPC KKMP NOLQ"
    Else
            ActiveCell.Offset(0, 7).Value = "LIB  IMPR CAPC KKMP NOLQ"
    End If

    UltCellProcess.Select
    ActiveCell.Offset(1, 0).Select
Loop

Set UltCellProcess = Nothing


'------------------------------------------------------------------
'------------------> Limpar justif. de OM no prazo <---------------

' Definir a tabela como um objeto
Dim tRelatorio As ListObject
Set tRelatorio = ActiveSheet.ListObjects("TabelaRelatorio")

For i = 1 To tRelatorio.ListColumns(16).DataBodyRange.Rows.Count
    tRelatorio.DataBodyRange(i, 16).Select
    
    If ActiveCell.Offset(0, 3).Value = "Realizado, dentro do prazo" Or ActiveCell.Offset(0, 3).Value = "Em aberto, dentro do prazo" Then
        ActiveCell.ClearContents
    Else
    End If
Next


'------------------------------------------------------------------
'------------------> Excluir linhas de OM scan <-------------------

Dim Linha As Long

With Planilha5
    For Linha = .Cells(.Rows.Count, "N").End(xlUp).Row To 6 Step -1
        If .Cells(Linha, "N") = "Cancelado" Then
            .Rows(Linha).Delete
        End If
    Next Linha
End With


'------------------------------------------------------------------
'------------------> Ajustar altura das linhas <-------------------
'Rows("6:6").Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.RowHeight = 16


'------------------> Finaliza com mensagem
MsgBox ("Relatório de Preventivas Exportado")

Else
'------------------------------------------------------------------
'------------------> Caso o usuário marque AnswerNo <--------------
MsgBox ("Relatório interrompido")
End If

' Msg de tratamento de erros, encerrando a sub e todo relatorio
Exit Sub
msgError: MsgBox "Não há soliciações para incuir ao relatório."

carteira1.HeaderRowRange.AutoFilter
carteira1.HeaderRowRange.AutoFilter
ActiveSheet.Range("A6").Select
Application.ScreenUpdating = True
End Sub











Sub RelatorioCorretivas_cart1()
'
' Macro para gerar relatorio

Dim ws As Excel.Worksheet
Dim carteira1 As ListObject
Set carteira1 = ActiveSheet.ListObjects("Tabela1")

'------------------> O processo só iniciará caso a AnswerYes = True
Dim AnswerYes As String
Dim AnswerNo As String

AnswerYes = MsgBox("Você já confirmou todas as Ordens/Notas executadas dentro deste período? (Ordens confirmadas após o dia 15, deverá ter status LIB)", vbYesNo, "Verificação")
If AnswerYes = vbYes Then
'--v--v--v--v--v--v--v--v-- Executar codigo abaixo. Completo!
Application.ScreenUpdating = False
Application.CutCopyMode = False


' START
Sheets("Relatório").Select ' <---
'------------------> Fazer limpeza da planilha de relatório

    ' Primeiro limparemos o conteúdo
    Range("TabelaRelatorio").Select
    Selection.ClearContents


Sheets("01-Carteira").Select ' <---
'------------------> Fazer filtros e critérios necessário

' Limpar possíveis filtros deixados pelo usuário antes do processo
carteira1.HeaderRowRange.AutoFilter
carteira1.HeaderRowRange.AutoFilter

    ' Filtrando as preventivas
    Range("Tabela1").AutoFilter Field:=10, Criteria1:="*Corretiva*"
    
    ' Filtrando (Field 5 = Data de Liberação)
    Range("Tabela1").AutoFilter Field:=5, Criteria1:= _
    "<=" & Format(Planilha0.Range("DataFinal").Value, "mm/dd/yyyy hh:mm")


'------------------> Transferir dados da Carteira, para o Relatório

Range("Tabela1").Select ' Seleciona a tabela
On Error Goto msgError ' Caso não haja linhas para copiar, salta a linha
Selection.SpecialCells(xlCellTypeVisible).Copy ' Copiando...

    
Sheets("Relatório").Select ' <---
'------------------> Voltando à planilha de relatório
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


'------------------------------------------------------------------
'------------------> Alterando STATUS das ordens <-----------------

' Alterar STATUS de outras corretivas
Dim w       As Worksheet
Dim UltCellProcess  As Range

Set w = Sheets("Relatório")
w.Range("A6").Select

Do While ActiveCell.Value <> ""
    Set UltCellProcess = ActiveCell
    If ActiveCell.Offset(0, 6).Value <> "" Then
            ActiveCell.Offset(0, 7).Value = "Atendimento Concluido"
    Else
            ActiveCell.Offset(0, 7).Value = "Assumida"
    End If

    UltCellProcess.Select
    ActiveCell.Offset(1, 0).Select
Loop

Set UltCellProcess = Nothing


' Alterar STATUS de corretivas do SAP
Dim tRelatorio As ListObject
Set tRelatorio = ActiveSheet.ListObjects("TabelaRelatorio")

' Loop por todas as linhas da coluna H da tabela
For i = 1 To tRelatorio.ListColumns(8).DataBodyRange.Rows.Count
    tRelatorio.DataBodyRange(i, 8).Select
        
    If ActiveCell.Offset(0, -1).Value <> "" And ActiveCell.Offset(0, 2).Value = "SAP Corretiva" Then
        ActiveCell.Value = "ENTE CONF IMPR CAPC KKMP NOLQ"
    Else
        If ActiveCell.Offset(0, -1).Value = "" And ActiveCell.Offset(0, 2).Value = "SAP Corretiva" Then
            ActiveCell.Value = "LIB  IMPR CAPC KKMP NOLQ"
        Else
        End If
    End If
Next i


'------------------------------------------------------------------
'------------------> Formula de contar dias <----------------------
For i = 1 To tRelatorio.ListColumns(12).DataBodyRange.Rows.Count
    tRelatorio.DataBodyRange(i, 12).Select
    
    ActiveCell.FormulaR1C1 = _
        "=IF([@[Data Conclusão]]<>"""",[@[Data Vencimento]]-[@[Data Conclusão]],[@[Data Vencimento]]-DataFinal)"
Next


'------------------> Arredondando dias abaixo de zero <------------
For i = 1 To tRelatorio.ListColumns(12).DataBodyRange.Rows.Count
    tRelatorio.DataBodyRange(i, 12).Select

    If ActiveCell.Value < 1 And ActiveCell.Value > -1 Then
        ActiveCell.Value = 0
    Else
        ActiveCell.Value = ActiveCell
    End If
Next


'------------------> Formula de status de medição <----------------
For i = 1 To tRelatorio.ListColumns(19).DataBodyRange.Rows.Count
    tRelatorio.DataBodyRange(i, 19).Select
    
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(IF([@[Data Conclusão]]<>"""",""Realizado, "",""Em aberto, ""),IF([@Atraso]>=0,""dentro do prazo"",""fora do prazo""),IF(AND([@Atraso]<0,[@Justificativa]<>""""),"" abonado"",""""))"
Next


'------------------------------------------------------------------
'------------------> Limpar justif. de OM no prazo <---------------

' A tabela ja foi definica como objeto la em cima
For i = 1 To tRelatorio.ListColumns(16).DataBodyRange.Rows.Count
    tRelatorio.DataBodyRange(i, 16).Select
    
    If ActiveCell.Offset(0, 3).Value = "Realizado, dentro do prazo" Or ActiveCell.Offset(0, 3).Value = "Em aberto, dentro do prazo" Then
        ActiveCell.ClearContents
    Else
    End If
Next


'------------------------------------------------------------------
'------------------> Excluir linhas de OM scan <-------------------

Dim Linha As Long

With Planilha5
    For Linha = .Cells(.Rows.Count, "N").End(xlUp).Row To 6 Step -1
        If .Cells(Linha, "N") = "Cancelado" Then
            .Rows(Linha).Delete
        End If
    Next Linha
End With


'------------------------------------------------------------------
'------------------> Ajustar altura das linhas <-------------------
'Rows("6:6").Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.RowHeight = 16


'------------------> Finaliza com mensagem
MsgBox ("Relatório de Preventivas Exportado")

Else
'------------------------------------------------------------------
'------------------> Caso o usuário marque AnswerNo <--------------
MsgBox ("Relatório interrompido")
End If

' Msg de tratamento de erros, encerrando a sub e todo relatorio
Exit Sub
msgError: MsgBox "Não há soliciações para incuir ao relatório."

carteira1.HeaderRowRange.AutoFilter
carteira1.HeaderRowRange.AutoFilter
ActiveSheet.Range("A6").Select
Application.ScreenUpdating = True
End Sub