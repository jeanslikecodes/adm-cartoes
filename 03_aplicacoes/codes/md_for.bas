Attribute VB_Name = "md_for"
' os métodos limpam e formatam as informações do pdf

Sub extract_and_transform()

    line_classification
    'move_to_clean_base

End Sub

Sub line_classification()

    Sheets(shBO).Select
    frBO = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    For rwBO = 2 To frBO
        vConteudo = LCase(Sheets(shBO).Range("B" & rwBO).Value & " ")
        vVarConteudo = Split(vConteudo, " ")
        
        vLP = UBound(vVarConteudo) ' last position of array vVarConteudo
    
        If vVarConteudo(0) = "cartão" Or vVarConteudo(0) = "natureza" Then
            Select Case vVarConteudo(0)
                Case "cartão": vClasse = "x"
                Case "natureza": vClasse = "y"
            End Select
        End If
        
        If Len(vVarConteudo(0)) = 2 And Len(vConteudo) > 4 Then
            If vVarConteudo(vLP) = "710-Caixa" Then
                vClasse = "w"
            Else
                vClasse = "z"
            End If
        End If
        
        If vClasse <> "" Then
            If vClasse = "w" Then
                For rwAux = rwBO To rwBO + 2
                    Sheets(shBO).Range("C" & rwAux).Value = vClasse
                Next
                
                rwBO = rwBO + 2
            Else
                Sheets(shBO).Range("C" & rwBO).Value = vClasse
            End If
        End If
        
        vClasse = ""
    Next
    
End Sub


