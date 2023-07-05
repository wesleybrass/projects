Sub Enviar_email()
    
Set Objeto_outlook = CreateObject("Outlook.application")
    
    For linha = 2 To Range("A1").End(xlDown).Row
        Set Email = Objeto_outlook.createitem(0)
        Email.display
        assinatura = Email.HtmlBody
            
        Email.To = Cells(linha, 1).Value
        Email.CC = "wesley.leite@tdconstrutora.com.br"
            
        Email.Subject = "Teste de VBA"
        Email.HtmlBody = "Boa tarde, " & Cells(linha, 2).Value & ". " & Cells(linha, 3).Value & assinatura
'"C:\Users\TD CONSTRUCOES\Documents\Apostilas Hastag Treinamentos\VBA - Integração VBA com Outlook\Relatórios Vendas\Vendas - Diego Amorim.xlsx"
        Email.attachments.Add (ThisWorkbook.Path & "\Relatórios Vendas\Vendas - " & Cells(linha, 2).Value & ".xlsx")
        Email.send
    Next

End Sub


'______INSTRUÇÕES DE USO__________________________________________

.Display
.To               '-> Para quem vai mandar
.CC               '-> Cópia
.BCC              '-> Cópia Oculta
.Subject          '-> Assunto
.HtmlBody         '-> Corpo do e-mail
.Attachments.Add  '-> Anexos


'inserir texto simples sem formatação com Html:
variavel.HtmlBody = "<p> & insira aqui o texto & "<\p>

'inserir texto com formatação com Html:
variavel.HtmlBody = "<p style=""font-size:15px"">" & insira aqui o texto & "<\p>"

'Outras Opções de formatação: font-family:calibri;color:black;font-weight:bold
variavel.HtmlBody = "<p style=""font-size:15px;font-family:calibri;color:black;font-weight:bold"">" & insira aqui o texto & "<\p>"

End Sub