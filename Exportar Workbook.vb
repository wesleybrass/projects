Sub ExportarWorkbook()
	
	'Declaração de variáveis
    Dim LoclArqv As Range
    Dim NameArqv As Range
    
	'Definindo celulas com a informação necessária
    Set LoclArqv = Sheets("planilha1").Range("A1")
    Set NameArqv = Sheets("planilha1").Range("B2)
    
	'Condicional
    If NameArqv <> False Then
        ActiveWorkbook.SaveCopyAs LoclArqv & NameArqv
    End If
    
End Sub


