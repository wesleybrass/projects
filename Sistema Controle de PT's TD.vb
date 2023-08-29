Private Sub btn_procurar_Click()
    '___Atualiza a caixa de listagem
    Call atualiza_lb_lista
End Sub
-----------------------------------------------------
Private Sub btn_limpar_Click()
    '___Limpa os campos
    tb_id.Value = ""
    tb_pt.Value = ""
    tb_ss.Value = ""
    cb_contrato.Value = ""
    tb_descricao.Value = ""
    tb_data.Value = Format(Date, "dd/mm/yyyy")
    tb_local.Value = ""
    tb_observacoes.Value = ""
End Sub
-----------------------------------------------------
Private Sub btn_saveAll_Click()
    ' Salva a planilha
    ThisWorkbook.Save
    MsgBox ("Planilha salva com sucesso")
End Sub
-----------------------------------------------------
Private Sub cb_contrato_Change()
    '___Atualiza as SS's
    Call carrega_SSs
End Sub
-----------------------------------------------------
Private Sub rb_aditivo_Click()
    '___Atualiza a caixa de listagem
    Call atualiza_lb_lista
End Sub
-----------------------------------------------------
Private Sub rb_tampao_Click()
    '___Atualiza a caixa de listagem
    Call atualiza_lb_lista
End Sub
-----------------------------------------------------
Private Sub rb_global_Click()
    '___Atualiza a caixa de listagem
    Call atualiza_lb_lista
End Sub
-----------------------------------------------------
Private Sub rb_verTudo_Click()
    '___Atualiza a caixa de listagem
    Call atualiza_lb_lista
End Sub
-----------------------------------------------------
Private Sub lb_lista_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    ' Toda vez que um produto for clicado duas vezes, todas as caixas ficarão preenchidas com as informações dele
    tb_pt.Value = lb_lista.List(lb_lista.ListIndex, 1)
    cb_contrato.Value = lb_lista.List(lb_lista.ListIndex, 2)
    tb_ss.Value = lb_lista.List(lb_lista.ListIndex, 3)
    tb_descricao.Value = lb_lista.List(lb_lista.ListIndex, 4)
    tb_data.Value = CDate(lb_lista.List(lb_lista.ListIndex, 5))
    tb_local.Value = lb_lista.List(lb_lista.ListIndex, 6)
    tb_observacoes.Value = lb_lista.List(lb_lista.ListIndex, 7)
    
    tb_id.Value = lb_lista.List(lb_lista.ListIndex, 0)

End Sub
-----------------------------------------------------
'___ Incluir registros
Private Sub btn_incluir_Click()

    limiteSS = 13
    limiteData = 10
    TamanhoSS = Len(tb_ss.Value)
    TamanhoData = Len(tb_data.Value)
    
    If tb_pt.Value = "" Then
        MsgBox ("Preencha o número da PT")
        Exit Sub
    End If
    
    If TamanhoSS <> limiteSS Then
        MsgBox ("Está faltando caracters no campo da SS")
        Exit Sub
    End If
    
    If tb_descricao.Value = "" Then
        MsgBox ("Preencha uma descrição para a PT")
        Exit Sub
    End If
    
    If cb_contrato.Value = "" Then
        MsgBox ("Preencha a qual contrato pertence")
        Exit Sub
    End If
    
    If TamanhoData <> limiteData Then
        MsgBox ("Preencha a data no formato: dd/mm/aaaa")
        Exit Sub
    End If
    
    '___ Descobrindo a ultima linha
    Linha = Sheets("PTs").Range("B1000000").End(xlUp).Row + 1
    
    If Linha = 2 Then
        Sheets("PTs").Cells(Linha, 1) = 1
    Else
        Sheets("PTs").Cells(Linha, 1) = WorksheetFunction.Max(Sheets("PTs").Range("A:A")) + 1
    End If
    
    Sheets("PTs").Cells(Linha, 2).Value = tb_pt.Value
    Sheets("PTs").Cells(Linha, 3).Value = cb_contrato.Value
    Sheets("PTs").Cells(Linha, 4).Value = tb_ss.Value
    Sheets("PTs").Cells(Linha, 5).Value = tb_descricao.Value
    Sheets("PTs").Cells(Linha, 6).Value = CDate(tb_data.Value)
    Sheets("PTs").Cells(Linha, 7).Value = tb_local.Value
    Sheets("PTs").Cells(Linha, 8).Value = tb_observacoes.Value
    
    tb_id.Value = ""
    tb_pt.Value = ""
    tb_ss.Value = ""
    cb_contrato.Value = ""
    tb_descricao.Value = ""
    tb_data.Value = Format(Date, "dd/mm/yyyy")
    tb_local.Value = ""
    tb_observacoes.Value = ""
    
    Call atualiza_lb_lista
    Call carrega_SSs
    MsgBox ("Transação adicionada com sucesso")
End Sub
-----------------------------------------------------
'___ Excluir registros
Private Sub btn_excluir_Click()

    If tb_pt.Value = "" Then
        MsgBox ("Cique duas vezes na linha que deseja excluir")
        Exit Sub
    End If
    
    Linha = Sheets("PTs").Range("A:A").Find(tb_id.Value).Row
    Sheets("PTs").Range(Linha & ":" & Linha).Delete Shift:=xlUp
    
    tb_id.Value = ""
    tb_pt.Value = ""
    tb_ss.Value = ""
    cb_contrato.Value = ""
    tb_descricao.Value = ""
    tb_data.Value = Format(Date, "dd/mm/yyyy")
    tb_local.Value = ""
    tb_observacoes.Value = ""
    
    Call atualiza_lb_lista
    Call carrega_SSs
    MsgBox ("Transação excluída com sucesso")
End Sub
-----------------------------------------------------
'
'
'
'
'
'////////////////////////////////////////////////////////////////////////////
'/////////////////// ABERTURA E FECHAMENTO DO FORMS /////////////////////////
'////////////////////////////////////////////////////////////////////////////

'___Executado quando abrir o formulário
Private Sub UserForm_Initialize()

Sheets("PTs").Unprotect Password:="0000"
Sheets("Inicio").Unprotect Password:="0000"
Sheets("Inicio").Activate
    
    '___Desligando ferramentas e recursos
    Application.DisplayFullScreen = True
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFormulaBar = False
        
    '___Inicializa a caixa de data com a data de hoje
    tb_data.Value = Format(Date, "dd/mm/yyyy")
    tb_dataInicio.Value = Format(Date, "dd/mm/yyyy")
    tb_dataFim.Value = Format(Date, "dd/mm/yyyy")
    
    '___Carrega as informações para as caixas
    cb_contrato.AddItem "4600672973 (Tampão)"
    cb_contrato.AddItem "4600667911 (Aditivo)"
    cb_contrato.AddItem "4600674820 (Global)"
    rb_verTudo.Value = True
    
    Sheets("Inicio").Activate
    Application.Wait Now() + TimeValue("00:00:01") ' Pause...
    Application.ScreenUpdating = False
    
    '___Carrega as informações na comboBox de SS's e na Caixa de listagem
    Call atualiza_lb_lista
    Call carrega_SSs

End Sub
-----------------------------------------------------
'___Executado quando fechar o formulário
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    '___Ativa as linhas de grade, exibe a barra de fórmulas, as abas, os títulos e tira o Excel da tela cheia
    Sheets("Inicio").Activate
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    
    '___Ativa a atualização de tela
    Application.ScreenUpdating = True
    
    Sheets("Impressao").Cells.Clear
    Sheets("Lista").Cells.Clear

Sheets("PTs").Protect Password:="0000"
Sheets("Inicio").Protect Password:="0000"
End Sub
-----------------------------------------------------
'
'
'
'
'
'////////////////////////////////////////////////////////////////////////////
'///////////////////////////// MACROS E FUNCTIONS ///////////////////////////
'////////////////////////////////////////////////////////////////////////////

Sub carrega_SSs()
    ' Descobre a última linha da aba de Premissas e carrega as informações para a caixa de SS's
    If cb_contrato.Value = "" Then
        Exit Sub
    End If
    
    If cb_contrato.Text = "4600667911 (Aditivo)" Then
        Linha = Sheets("Premissas").Range("A1000000").End(xlUp).Row + 1
        If Linha = 1 Then Linha = 2
        tb_ss.RowSource = "Premissas!A2:A" & Linha
    ElseIf cb_contrato.Text = "4600672973 (Tampão)" Then
        Linha = Sheets("Premissas").Range("C1000000").End(xlUp).Row + 1
        If Linha = 1 Then Linha = 2
        tb_ss.RowSource = "Premissas!C2:C" & Linha
    ElseIf cb_contrato.Text = "4600674820 (Global)" Then
        Linha = Sheets("Premissas").Range("E1000000").End(xlUp).Row + 1
        If Linha = 1 Then Linha = 2
        tb_ss.RowSource = "Premissas!E2:E" & Linha
    End If
End Sub
-----------------------------------------------------
Sub atualiza_lb_lista()

Sheets("PTs").Unprotect Password:="0000"
    '___Limpa todos os filtros da base de dados
    Sheets("PTs").AutoFilterMode = False
    
    '___Filtra as informações na base de dados com o status selecionado no formulário
    If rb_aditivo = True Then
        Sheets("PTs").UsedRange.AutoFilter 3, "4600667911 (Aditivo)"
    ElseIf rb_tampao = True Then
        Sheets("PTs").UsedRange.AutoFilter 3, "4600672973 (Tampão)"
    ElseIf rb_global = True Then
        Sheets("PTs").UsedRange.AutoFilter 3, "4600674820 (Global)"
    End If
    
    '___Filtra as informações da PTs e cola na aba Listagem
    If tb_procurar.Value <> "" Then
        Sheets("PTs").UsedRange.AutoFilter 4, "*" & tb_procurar.Value & "*"
    End If
    
    '___Limpa as informações da aba Lista e cola as informações filtradas
    Sheets("Lista").UsedRange.Clear
    Sheets("PTs").UsedRange.Copy
    Sheets("Lista").Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = xlCopy
    
    '___Tira o filtro da base de dados principal
    Sheets("PTs").AutoFilterMode = False
    
    '___Mostrar as informações da aba criada na caixa de listagem
    ' Descobre a última linha preenchida da aba "Lista"
    Linha = Sheets("Lista").Range("B1000000").End(xlUp).Row
    ' Acertar a informação da linha se não tiver nenhuma informação preenchida na tabela
    If Linha = 1 Then Linha = 2
    
    '___Ordenar do mais novo ao mais antigo
    ActiveWorkbook.Worksheets("Lista").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lista").Sort.SortFields.Add2 Key:=Range _
        ("F2:F" & Linha), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Lista").Sort
        .SetRange Range("A2:H" & Linha)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    label_qtde.Caption = Sheets("Lista").Range("A1000000").End(xlUp).Row - 1
    
    
    ' Carrega a caixa de listagem com 8 colunas e todas as informações da aba Caixa_transações
    With FormPrincipal.lb_lista
        .ColumnCount = 8
        .ColumnHeads = True
        .ColumnWidths = "0;45;90;70;300;60;85;85"
        .RowSource = "Lista!A2:H" & Linha
    End With
End Sub
-----------------------------------------------------
'
'
'
'
'
'////////////////////////////////////////////////////////////////////////////
'///////////////////////////// IMPRESSÃO ////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////

Sub btn_gerarPDF_Click()
Dim Titulo As String
NumSS = tb_procurar.Value
    
    'Limpando tudo na aba impressão
    Sheets("Impressao").Cells.Clear
    DataINI = Format(tb_dataInicio.Value, "mm/dd/yyyy")
    DataFIM = Format(tb_dataFim.Value, "mm/dd/yyyy")
    Titulo = "Relação de PT's Cabiúnas - Período " & Format(DataINI, "dd/mm/yyyy") & " a " & Format(DataFIM, "dd/mm/yyyy")
    
    ' Escrevendo o titulo de estoque
    Sheets("Impressao").Range("A1").Value = Titulo
    Sheets("Impressao").Range("A1:G1").Merge
    Sheets("Impressao").Rows(1).Font.Bold = True
    Sheets("Impressao").Rows(3).Font.Bold = True
    Sheets("Impressao").Rows(1).Font.Size = 12
    
    ' Reaproveitando os filtros da lista
    Sheets("Lista").Activate
    Columns("A:A").Select
    Selection.Delete
    Application.CutCopyMode = False
    
    ' Filtro de datas/periodos
    Sheets("Lista").Activate
    ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:= _
        ">=" & DataINI, Operator:=xlAnd, Criteria2:="<=" & DataFIM
    
    ' Copia de: Lista para Impressao
    Sheets("Lista").Range("A1").CurrentRegion.Copy Sheets("Impressao").Range("A3")
    Application.CutCopyMode = False
    
    ' Ajustando alinhamentos, fonte e colunas
    Sheets("Impressao").Activate
    ultLinha = Sheets("Impressao").Range("A1000000").End(xlUp).Row
    Cells.Select
    Selection.Font.Size = 9
    Sheets("Impressao").Rows(1).Font.Size = 12
    Sheets("Impressao").Rows(3).Font.Bold = True
    Sheets("Impressao").Rows.HorizontalAlignment = xlHAlignCenter
    Sheets("Impressao").Rows.VerticalAlignment = xlHAlignCenter
    Sheets("Impressao").Columns("D:G").WrapText = True
    Rows("1:3").RowHeight = 20
    Sheets("Impressao").Range("A3:G" & ultLinha).Borders.ThemeColor = 3
    Sheets("Lista").AutoFilterMode = False

    Dim nomeArquivo As String
    localArquivo = ThisWorkbook.Path
    nomeArquivo = "\Relatórios exportados\SS-" & NumSS & " - Relação de PT's CAB - Período " & Format(DataINI, "dd-mm-yyyy") & " a " & Format(DataFIM, "dd-mm-yyyy")
    caminho = localArquivo + nomeArquivo

'Unload FormPrincipal

    Sheets("Impressao").ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminho, _
    Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False

End Sub

Sub btn_gerarEXCEL_Click()
Dim Titulo As String
NumSS = tb_procurar.Value
    
    'Limpando tudo na aba impressão
    Sheets("Impressao").Cells.Clear
    DataINI = Format(tb_dataInicio.Value, "mm/dd/yyyy")
    DataFIM = Format(tb_dataFim.Value, "mm/dd/yyyy")
    Titulo = "Relação de PT's Cabiúnas - Período " & Format(DataINI, "dd/mm/yyyy") & " a " & Format(DataFIM, "dd/mm/yyyy")
    
    ' Escrevendo o titulo de estoque
    Sheets("Impressao").Range("A1").Value = Titulo
    Sheets("Impressao").Range("A1:G1").Merge
    Sheets("Impressao").Rows(1).Font.Bold = True
    Sheets("Impressao").Rows(3).Font.Bold = True
    Sheets("Impressao").Rows(1).Font.Size = 12
    
    ' Reaproveitando os filtros da lista
    Sheets("Lista").Activate
    Columns("A:A").Select
    Selection.Delete
    Application.CutCopyMode = False
    
    ' Filtro de datas/periodos
    Sheets("Lista").Activate
    ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:= _
        ">=" & DataINI, Operator:=xlAnd, Criteria2:="<=" & DataFIM
    
    ' Copia de: Lista para Impressao
    Sheets("Lista").Range("A1").CurrentRegion.Copy Sheets("Impressao").Range("A3")
    Application.CutCopyMode = False
    
    ' Ajustando alinhamentos, fonte e colunas
    Sheets("Impressao").Activate
    ultLinha = Sheets("Impressao").Range("A1000000").End(xlUp).Row
    Cells.Select
    Selection.Font.Size = 9
    Sheets("Impressao").Rows(1).Font.Size = 12
    Sheets("Impressao").Rows(3).Font.Bold = True
    Sheets("Impressao").Rows.HorizontalAlignment = xlHAlignCenter
    Sheets("Impressao").Rows.VerticalAlignment = xlHAlignCenter
    Sheets("Impressao").Columns("D:G").WrapText = True
    Rows("1:3").RowHeight = 20
    Sheets("Impressao").Range("A3:G" & ultLinha).Borders.ThemeColor = 3
    Sheets("Lista").AutoFilterMode = False

    Dim nomeArquivo As String
    localArquivo = ThisWorkbook.Path
    nomeArquivo = "\Relatórios exportados\SS-" & NumSS & " - Relação de PT's CAB - Período " & Format(DataINI, "dd-mm-yyyy") & " a " & Format(DataFIM, "dd-mm-yyyy")
    caminho = localArquivo + nomeArquivo

Unload FormPrincipal

    Worksheets("Impressao").Copy
    With ActiveWorkbook
        .SaveAs Filename:=caminho, FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
    End With
    Workbooks.Open caminho
    
End Sub
-----------------------------------------------------











---------------------------------MACROS--------------
Sub mostra_formulario()
    FormPrincipal.Show
End Sub
Sub mostra_ajuda()
    Sheets("Inicio").Activate
    Ajuda.Show
End Sub

Sub limpar()
    On Error Goto sair
    ActiveSheet.Unprotect Password:="0000"
    Sheets("Inicio").Range("G8:N110").ClearContents
    Sheets("Inicio").Range("G8:N110").Interior.ColorIndex = 0
    Sheets("Inicio").Range("G8:N110").Font.ColorIndex = 1
    Range("H8").Select
    ActiveSheet.Protect Password:="0000"
sair:
    ActiveSheet.Protect Password:="0000"
End Sub

Sub colar()
    On Error Resume Next
    Sheets("Inicio").Range("H8").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
End Sub

Sub ColarClipBoard()
    Dim destinoPlanilha As Worksheet
    Dim clipboardData As DataObject
    
    Set clipboardData = New DataObject
    clipboardData.GetFromClipboard
    Set destinoPlanilha = ThisWorkbook.Sheets("Inicio")
    
    On Error Goto sair
    ActiveSheet.Unprotect Password:="0000"
    destinoPlanilha.Select
    destinoPlanilha.Range("H8").Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Protect Password:="0000"
sair:
    ActiveSheet.Protect Password:="0000"
End Sub

Sub AlimentarPlanilha()
    
    Application.ScreenUpdating = False
    Sheets("Inicio").Unprotect Password:="0000"
    Sheets("PTs").Unprotect Password:="0000"
    Sheets("Inicio").Activate
    
    '___ Verificando se as primeiras celulas estao vazias
    If Range("K8").Value = "" Or Range("L8").Value = "" Or Range("H8").Value = "" And Range("J8").Value = "" Then
        Sheets("Inicio").Protect Password:="0000"
        Sheets("PTs").Protect Password:="0000"
        Application.ScreenUpdating = True
        MsgBox ("Não há nenhuma informação na planilha para ser transferida")
        Exit Sub
    End If
    
'-------------------------
'--> Validação de dados: Check Contrato
    Dim Linha As Long
    With Planilha1
        For Linha = .Cells(.Rows.Count, "H").End(xlUp).Row To 8 Step -1
            If .Cells(Linha, "I").Value = "4600672973 (Tampão)" Or .Cells(Linha, "I").Value = "4600667911 (Aditivo)" Or .Cells(Linha, "I").Value = "4600674820 (Global)" Then
                .Cells(Linha, "G").Value = "ok"
                .Cells(Linha, "I").Font.ColorIndex = 1
            Else
                .Cells(Linha, "G").Value = "erro"
                .Cells(Linha, "I").Font.ColorIndex = 3
            End If
        Next Linha
    End With
'-------------------------
'--> Validação de dados: Check SS's
    Dim Linha2 As Long
    With Planilha1
        For Linha2 = .Cells(.Rows.Count, "H").End(xlUp).Row To 8 Step -1
            If Len(.Cells(Linha2, "J").Value) <> 13 Then
                .Cells(Linha2, "G").Value = "erro"
                .Cells(Linha2, "J").Font.ColorIndex = 3
            Else
                .Cells(Linha2, "J").Font.ColorIndex = 1
                If .Cells(Linha2, "G").Value = "ok" Then .Cells(Linha2, "G").Value = "ok"
                If .Cells(Linha2, "G").Value = "erro" Then .Cells(Linha2, "G").Value = "erro"
            End If
        Next Linha2
    End With
'-------------------------
'--> Validação de dados: Sem numero de PT
    Dim Linha3 As Long
    With Planilha1
        For Linha3 = .Cells(.Rows.Count, "J").End(xlUp).Row To 8 Step -1
            If .Cells(Linha3, "H").Value = "" Then
                .Cells(Linha3, "G").Value = "PT?"
                .Cells(Linha3, "H").Interior.ColorIndex = 3
            Else
                .Cells(Linha3, "H").Interior.ColorIndex = 0
                If .Cells(Linha3, "G").Value = "ok" Then .Cells(Linha3, "G").Value = "ok"
                If .Cells(Linha3, "G").Value = "erro" Then .Cells(Linha3, "G").Value = "erro"
            End If
        Next Linha3
    End With
'-------------------------
'--> Se houver alguma linha com erro, sai da Sub
    With Planilha1
        For Linha = .Cells(.Rows.Count, "K").End(xlUp).Row To 8 Step -1
            If .Cells(Linha, "G").Value <> "ok" Then
                MsgBox ("Corrija as linhas com informações incorretas")
                Sheets("Inicio").Protect Password:="0000"
                Sheets("PTs").Protect Password:="0000"
                Application.ScreenUpdating = True
                Exit Sub
            End If
        Next Linha
    End With
'-------------------------
    
    Sheets("Inicio").Activate
    UltLn_Inicio = Sheets("Inicio").Range("H110").End(xlUp).Row
    If UltLn_Inicio = 7 Then UltLn_Inicio = 8
    
    'Iniciando processo de copiar e colar
    Range("H8:N" & UltLn_Inicio).Select
    Selection.Copy
    UltLn_PTs = Sheets("PTs").Range("B500000").End(xlUp).Row + 1
    Sheets("PTs").Range("B" & UltLn_PTs).PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    
    Sheets("PTs").Activate
    UltLn_PTs = Sheets("PTs").Range("B500000").End(xlUp).Row
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A" & UltLn_PTs), Type:=xlFillSeries
    
    Sheets("Inicio").Activate
    Range("G8:N" & UltLn_Inicio).ClearContents
    Range("G8:N" & UltLn_Inicio).Interior.ColorIndex = 0
    Range("G8:N" & UltLn_Inicio).Font.ColorIndex = 1
    Range("H8").Select
    
    ActiveSheet.Protect Password:="0000"
    Sheets("PTs").Protect Password:="0000"
    
    Application.ScreenUpdating = True
    MsgBox ("Alimentação concluída")
End Sub
-----------------------------------------------------
Private Sub Workbook_Open()

    ' Maximiza o Excel
    Application.WindowState = xlMaximized
    Sheets("Inicio").Activate
    ActiveWorkbook.Protect Password:="0000", Structure:=True, Windows:=False
    ActiveSheet.Protect Password:="0000"

    Application.DisplayFullScreen = True
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFormulaBar = False
    Range("H8").Select
    
    ' Pause...
    Application.Wait Now() + TimeValue("00:00:01")

    ' Desativa a atualização de tela
    Application.ScreenUpdating = False
    
    ' Exibe o Formulário de Estoque
    FormPrincipal.Show
End Sub

