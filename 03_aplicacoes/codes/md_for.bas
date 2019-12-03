Attribute VB_Name = "md_for"
' os métodos limpam e formatam as informações do pdf

Sub extract_and_transform()

    line_classification
    move_to_clean_base
    clean_information
    id_bl
    move_to_consolidated_base

End Sub

Sub line_classification()
    
    Sheets(shBO).Select
    frBO = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    For rwBO = 2 To frBO
        vConteudo = LCase(Sheets(shBO).Range("B" & rwBO).Value & " ")
        vVarconteudo = Split(vConteudo, " ")
        
        vLP = UBound(vVarconteudo) ' last position of array vVarConteudo
    
        If vVarconteudo(0) = "cartão" Or vVarconteudo(0) = "natureza" Then
            Select Case vVarconteudo(0)
                Case "cartão": vClasse = "x"
                Case "natureza": vClasse = "y"
            End Select
        End If
        
        If Len(vVarconteudo(0)) = 2 And Len(vConteudo) > 4 Then
            If vVarconteudo(vLP) = "710-Caixa" Then
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

Sub clean_information()
    
    Sheets(shBL).Select
    frBL = Sheets(shBL).Cells(Rows.Count, 1).End(xlUp).Row
            
    For rwBL = 2 To frBL
        vClasse = Sheets(shBL).Range("D" & rwBL).Value
        
        ' vConteudoB
        vConteudoB = Sheets(shBL).Range("B" & rwBL).Value

        vGmb = Mid(vConteudoB, InStr(1, vConteudoB, "_") + 1, 3)
        vAnoMes = Mid(vConteudoB, 1, InStr(1, vConteudoB, "_") - 1)
        
        'vConteudo C
        vConteudoC = Sheets(shBL).Range("C" & rwBL).Value
        vVarConteudoC = Split(vConteudoC, " ")
        vLPC = UBound(vVarConteudoC) ' Last position conteudo C
        
        Select Case vClasse
            Case "x":
                vOperadora = LCase(Trim(vVarConteudoC(vLPC)))
                vBandeira = Replace(LCase(Trim(vVarConteudoC(3))), "american", "american express")
                
                Select Case vBandeira
                    Case "american express", "master": vMetodo = "credito"
                    Case "maestro": vMetodo = "debito"
                    Case "visa", "elo":
                        If LCase(Trim(vVarConteudoC(4))) = "electron" Or _
                            LCase(Trim(vVarConteudoC(4))) = "debito" Then
                            vMetodo = "debito"
                        Else
                            vMetodo = "credito"
                        End If
                End Select
            Case "y":
                vNatureza = LCase(Trim(Replace(vConteudoC, "Natureza Receita:", "")))
            Case "z", "w":
                vEmpresa = vVarConteudoC(0)
                vData = CDate(Left(vVarConteudoC(1), 2) & "/" & _
                            Mid(vVarConteudoC(1), 4, 2) & "/" & _
                            Right(vVarConteudoC(1), 4))
                
                If InStr(1, vVarConteudoC(vLPC - 6), "4-", vbTextCompare) > 0 Then
                    vVarNome = 7
                    vCC = LCase(Trim(vVarConteudoC(vLPC - 6)))
                Else
                    vVarNome = 6
                    vCC = LCase(Trim(vVarConteudoC(vLPC - 5)))
                End If
                
                vNome = ""
                
                For vV = 2 To vLPC - vVarNome
                    vNome = vNome & vVarConteudoC(vV) & " "
                Next
                
                vNome = Trim(Replace(vNome, " - ", ""))
                
                vLote = vVarConteudoC(vLPC - 4)
                vParcelas = vVarConteudoC(vLPC - 3)
                vParcelaBruta = CDbl(vVarConteudoC(vLPC - 2))
                vParcelaTaxa = CDbl(vVarConteudoC(vLPC - 1))
                vParcelaLiquida = CDbl(vVarConteudoC(vLPC))


                Sheets(shBL).Range("F" & rwBL).Value = vGmb
                Sheets(shBL).Range("G" & rwBL).Value = vEmpresa
                Sheets(shBL).Range("H" & rwBL).Value = vAnoMes
                Sheets(shBL).Range("I" & rwBL).Value = vData
                Sheets(shBL).Range("J" & rwBL).Value = vNatureza
                Sheets(shBL).Range("K" & rwBL).Value = vOperadora
                Sheets(shBL).Range("L" & rwBL).Value = vBandeira
                Sheets(shBL).Range("M" & rwBL).Value = vMetodo
                Sheets(shBL).Range("N" & rwBL).Value = vNome
                Sheets(shBL).Range("O" & rwBL).Value = vCC
                Sheets(shBL).Range("P" & rwBL).Value = vLote
                Sheets(shBL).Range("Q" & rwBL).Value = vParcelas
                Sheets(shBL).Range("R" & rwBL).Value = vParcelaBruta
                Sheets(shBL).Range("S" & rwBL).Value = vParcelaTaxa
                Sheets(shBL).Range("T" & rwBL).Value = vParcelaLiquida
        End Select
        
    Next
 
End Sub

Sub id_bl()
    
    Sheets(shBL).Select
    frBL = Sheets(shBL).Cells(Rows.Count, 1).End(xlUp).Row
    
    For rwBL = 2 To frBL
        If Sheets(shBL).Range("F" & rwBL).Value <> "" Then
             vGmb = Sheets(shBL).Range("F" & rwBL).Value
             vEmpresa = Sheets(shBL).Range("G" & rwBL).Value
             vAnoMes = Sheets(shBL).Range("H" & rwBL).Value
             vData = Sheets(shBL).Range("I" & rwBL).Value
             vNatureza = Sheets(shBL).Range("J" & rwBL).Value
             vOperadora = Sheets(shBL).Range("K" & rwBL).Value
             vBandeira = Sheets(shBL).Range("L" & rwBL).Value
             vMetodo = Sheets(shBL).Range("M" & rwBL).Value
             vNome = Sheets(shBL).Range("N" & rwBL).Value
             vParcelas = Sheets(shBL).Range("Q" & rwBL).Value
             vParcelaBruta = Format(Sheets(shBL).Range("R" & rwBL).Value, "#.#0")
             vParcelaBruta = Left(vParcelaBruta, InStr(1, vParcelaBruta, ",", vbTextCompare) + 1)
             
             vIdBL = "|" & vGmb & "|" & vEmpresa & "|" & vAnoMes & "|" & vData & "|" & _
                        vNatureza & "|" & vOperadora & "|" & vBandeira & "|" & vMetodo & "|" & _
                        vNome & "|" & vParcelas & "|" & vParcelaBruta & "|"
             
             Sheets(shBL).Range("E" & rwBL).Value = vIdBL
        End If
    Next
    
End Sub

Sub move_to_consolidated_base()
    
    Sheets(shBL).Select
    frBL = Sheets(shBL).Cells(Rows.Count, 1).End(xlUp).Row
    
    For rwBL = 2 To frBL
        vIdBL = Sheets(shBL).Range("E" & rwBL).Value
        
        If vIdBL <> "" Then
            vParcelas = Sheets(shBL).Range("Q" & rwBL).Value
        
            If vParcelas = 1 Then
                Sheets(shBL).Range("E" & rwBL & ":U" & rwBL).Select
                Selection.Copy
                
                Sheets(shBC).Select
                frBC = Sheets(shBC).Cells(Rows.Count, 1).End(xlUp).Row
            
                Sheets(shBC).Range("A" & frBC + 1).Select
                ActiveSheet.Paste
            Else
                On Error Resume Next
                rwBC = WorksheetFunction.Match(vIdBL, Sheets(shBC).Range("A:A"), 0)
            
                If rwBC = "" Then
                    Sheets(shBL).Range("E" & rwBL & ":U" & rwBL).Select
                    Selection.Copy
                
                    Sheets(shBC).Select
                    frBC = Sheets(shBC).Cells(Rows.Count, 1).End(xlUp).Row
            
                    Sheets(shBC).Range("A" & frBC + 1).Select
                    ActiveSheet.Paste
                Else
                    vParcelaBruta = Sheets(shBL).Range("R" & rwBL).Value
                    vParcelaTaxa = Sheets(shBL).Range("S" & rwBL).Value
                    vParcelaLiquida = Sheets(shBL).Range("T" & rwBL).Value
                    
                    Sheets(shBC).Select
                    Sheets(shBC).Range("N" & rwBC).Value = Sheets(shBC).Range("N" & rwBC).Value + vParcelaBruta
                    Sheets(shBC).Range("O" & rwBC).Value = Sheets(shBC).Range("O" & rwBC).Value + vParcelaTaxa
                    Sheets(shBC).Range("P" & rwBC).Value = Sheets(shBC).Range("P" & rwBC).Value + vParcelaLiquida
                End If
            End If
            
        End If
        
        rwBC = ""
        
        Sheets(shBL).Select
    Next

End Sub

