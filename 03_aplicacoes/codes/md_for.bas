Attribute VB_Name = "md_for"
' os métodos limpam e formatam as informações do pdf

Sub extract_and_transform()

    line_classification
    move_to_clean_base

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

Sub move_to_clean_base()

    Sheets(shBO).Select
    frBO = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    For rwBO = 2 To frBO
        vClasse = Sheets(shBO).Range("C" & rwBO).Value
        
        If vClasse <> "" Then
            vArquivo = Sheets(shBO).Range("A" & rwBO).Value
            
            If vClasse <> "w" Then
                vConteudo = Sheets(shBO).Range("B" & rwBO).Value
            Else
                vConteudc = Sheets(shBO).Range("B" & rwBO).Value & _
                                Sheets(shBO).Range("B" & rwBO + 1).Value & " " & _
                                Sheets(shBO).Range("B" & rwBO + 2).Value
                rwBO = rwBO + 2
            End If
            
            Sheets(shBL).Select
            frBL = Sheets(shBL).Cells(Rows.Count, 1).End(xlUp).Row
            
            If Sheets(shBL).Range("A" & frBL).Value = "registro" Then
                Sheets(shBL).Range("A" & frBL + 1).Value = 1
            Else
                Sheets(shBL).Range("A" & frBL + 1).Value = Sheets(shBL).Range("A" & frBL).Value + 1
            End If
            
            Sheets(shBL).Range("B" & frBL + 1).Value = vArquivo
            Sheets(shBL).Range("C" & frBL + 1).Value = vConteudo
            Sheets(shBL).Range("D" & frBL + 1).Value = vClasse
        End If
    Next
End Sub


