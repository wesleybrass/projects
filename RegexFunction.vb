Function regex(strInput As String, matchPattern As String, Optional ByVal outputPattern As String = "$0") As Variant
    Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
    Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
    Dim replaceNumber As Integer

    With inputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = matchPattern
    End With

    With outputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "$(d+)"
    End With

    With outReplaceRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With

    Set inputMatches = inputRegexObj.Execute(strInput)

    If inputMatches.Count = 0 Then
        regex = False
    Else

        Set replaceMatches = outputRegexObj.Execute(outputPattern)
        For Each replaceMatch In replaceMatches
            
            replaceNumber = replaceMatch.SubMatches(0)
            outReplaceRegexObj.Pattern = "$" & replaceNumber


            If replaceNumber = 0 Then
                outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).Value)
            Else
                If replaceNumber > inputMatches(0).SubMatches.Count Then
                    regex = CVErr(xlErrValue)
                    Exit Function
                Else
                    outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
                End If
            End If
        Next
        regex = outputPattern

    End If

End Function


'-----------------------------------------------------------------

Sub extrairCPFRG()

	For linha = 2 To 5
	    Cells(linha, 2).Value = regex(Cells(linha, 1).Value, "[0-9]{3}.[0-9]{3}.[0-9]{3}-[0-9]{2}")
	    Cells(linha, 3).Value = regex(Cells(linha, 1).Value, "[0-9]{2}.[0-9]{3}.[0-9]{3}-[0-9]{1}")
	Next

End Sub