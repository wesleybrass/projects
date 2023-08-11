Sub ExportarWorkbook()
	
    Dim LoclArqv As Range
    Dim NameArqv As Range
    
	'Definindo celulas com a informação necessária
    Set LoclArqv = Sheets("planilha1").Range("A1")
    Set NameArqv = Sheets("planilha1").Range("B2")

    If NameArqv <> False Then
        ActiveWorkbook.SaveCopyAs LoclArqv & NameArqv
    End If
    
End Sub

Sub ExportarPDF()

    localArquivo = ThisWorkbook.Path
    nomeArquivo = "\nome_do_arquivo"
    caminho = localArquivo + nomeArquivo

    Sheets("Impressao").ExportAsFixedFormat Type:=xlTypePDF, fileName:=caminho, _
    Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=True
    
End Sub
