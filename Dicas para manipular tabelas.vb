TRABALHANDO COM TABELAS

'Usando a tabela como uma variável de objeto no VBA
Dim minhaTabela As ListObject
Set minhaTabela = ActiveSheet.ListObjects("table1")

'Selecionar a tabela inteira
ActiveSheet.ListObjects("table1").Range.Select

'Selecionar o cabeçalho da tabela
ActiveSheet.ListObjects("table1").HeaderRowRange.Select

'Selecionar todos os dados da tabela
ActiveSheet.ListObjects("table1").DataBodyRange.Select

'Selecionar a segunda coluna inteira da tabela
ActiveSheet.ListObjects("table1").ListColumns(2).Range.Select

'Selecionar somente o conteúdo da segunda coluna da tabela
ActiveSheet.ListObjects("table1").ListColumns(2).DataBodyRange.Select

'Selecionar a segunda linha inteira da tabela
ActiveSheet.ListObjects("table1").ListRows(2).Range.Select

'Selecionar a segunda linha da quarta coluna da tabela
ActiveSheet.ListObjects("table1").DataBodyRange(2, 4).Select

'Selecionar a linha de totais da tabela
ActiveSheet.ListObjects("table1").TotalsRowRange.Select
-----------------------------------------------------------
'Limpar todo o conteúdo de uma tabela VBA Excel
'E excluir As linhas da tabela junto. Full clear

ActiveSheet.ListObjects("tabel1").DataBodyRange.Delete 
 
