Sub Exportar_Relatorios()
    '--> Se a opção estiver ativa, chama a macro sozinha
        If Sheets("Relatório de Atividades").Range("D24").Value = "MENSAL" Then
            Call mensal
            Exit Sub
        End If
        
    '--> Se a opção estiver ativa, chama a macro sozinha
        If Sheets("Relatório de Atividades").Range("D24").Value = "QUINZENAL" Then
            Call quinzenal
            Exit Sub
        End If
        
    '--> Se a opção estiver ativa, chama a macro sozinha
        If Sheets("Relatório de Atividades").Range("D24").Value = "SEMANAL" Then
            Call semanal
            Exit Sub
        End If
        
    '--> Se a opção estiver vazia
        If Sheets("Relatório de Atividades").Range("D24").Value = "" Then
            MsgBox "Escolha o período do seu relatório"
            Exit Sub
        End If
    End Sub
    
    
    
    '////////////////////////////////
    '////////////////////////////////
    '////////////////////////////////
    Private Sub semanal()
        Dim LocArqv As Range
        Dim NameArqv As Range
        Dim dataAnterior As Date
        
        ActiveSheet.Unprotect
        Sheets("Mês de Medição - Período").Visible = False
        Sheets("Funções").Visible = False
        Sheets("Colaboradores").Visible = False
        Sheets("Todas as SS's").Visible = False
        Application.ScreenUpdating = False
        
        Set LocArqv = Sheets("Mês de Medição - Período").Range("M2")
        Set NameArqv = Sheets("Mês de Medição - Período").Range("J10")
        dataAnterior = Range("E28").Value
        
        'Processo de salvamento
        If NameArqv <> False Then
            ActiveWorkbook.SaveCopyAs LocArqv & NameArqv
        End If
        
        For i = 1 To 3
            Sheets("Mês de Medição - Período").Range("M6").Copy
            Sheets("Relatório de Atividades").Range("E28").PasteSpecial xlPasteValuesAndNumberFormats
            Range("A14:J19").Select
            
            'Processo de salvamento
            If NameArqv <> False Then
                ActiveWorkbook.SaveCopyAs LocArqv & NameArqv
            End If
        Next i
        
        'De volta como era antes
        Range("E28").Value = dataAnterior
        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                True, AllowFormattingCells:=True
        Range("A14:J19").Select
        Application.ScreenUpdating = True
    End Sub
    
    Private Sub quinzenal()
    
        Dim LocArqv As Range
        Dim NameArqv As Range
        Dim dataAnterior As Date
        
        ActiveSheet.Unprotect
        Sheets("Mês de Medição - Período").Visible = False
        Sheets("Funções").Visible = False
        Sheets("Colaboradores").Visible = False
        Sheets("Todas as SS's").Visible = False
        Application.ScreenUpdating = False
        
        Set LocArqv = Sheets("Mês de Medição - Período").Range("M2")
        Set NameArqv = Sheets("Mês de Medição - Período").Range("J10")
        dataAnterior = Range("E28").Value
        
        'Processo de salvamento
        If NameArqv <> False Then
            ActiveWorkbook.SaveCopyAs LocArqv & NameArqv
        End If
    
        Sheets("Mês de Medição - Período").Range("M6").Copy
        Sheets("Relatório de Atividades").Range("E28").PasteSpecial xlPasteValuesAndNumberFormats
        Range("A14:J19").Select
        
        'Processo de salvamento
        If NameArqv <> False Then
            ActiveWorkbook.SaveCopyAs LocArqv & NameArqv
        End If
    
        'De volta como era antes
        Range("E28").Value = dataAnterior
        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                True, AllowFormattingCells:=True
        Range("A14:J19").Select
        Application.ScreenUpdating = True
    End Sub
    
    Private Sub mensal()
    
        Dim LocArqv As Range
        Dim NameArqv As Range
    
        ActiveSheet.Unprotect
        Sheets("Mês de Medição - Período").Visible = False
        Sheets("Funções").Visible = False
        Sheets("Colaboradores").Visible = False
        Sheets("Todas as SS's").Visible = False
        Application.ScreenUpdating = False
        
        Set LocArqv = Sheets("Mês de Medição - Período").Range("M2")
        Set NameArqv = Sheets("Mês de Medição - Período").Range("J10")
        
        'Processo de salvamento
        If NameArqv <> False Then
            ActiveWorkbook.SaveCopyAs LocArqv & NameArqv
        End If
        
        'De volta como era antes
        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                True, AllowFormattingCells:=True
        Range("A14:J19").Select
        Application.ScreenUpdating = True
    End Sub
