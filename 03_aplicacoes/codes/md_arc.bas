Attribute VB_Name = "md_arc"
Function check_existence_arq_in_pc(nameFile As String)

    ' checa se o arquivo que irá ser lido consta no painel de contrrole
    
    Dim rwArc As String
    
    Sheets(shPC).Select
    
    On Error Resume Next
    rwArc = WorksheetFunction.Match(nameFile, Sheets(shPC).Range("B:B"), 0)
    
    check_existence_arq_in_pc = rwArc
    
End Function

Sub copy_content_pdf(nameFile As String, sizefile As Long)

    Dim adobeApp As String
    Dim adobeStart
    
    adobeFile = pdfPath & "/" & nameFile
            
    If CInt(Round(sizefile / 1024, 0)) <= 220 Then
        seconds = 40
    Else
        seconds = 60
    End If
    
    minutes = 0
            
    thisWorkbook.FollowHyperlink adobeFile
            
    Application.Wait (Now + TimeSerial(0, 0, 1))                        ' Abre o adobe e espera carregar
            
    SendKeys "^a"                                                       'Select All
    SendKeys "^c"                                                       'Copy
            
    Application.Wait (Now() + TimeSerial(0, minutes, seconds))
            
    SendKeys ("%{F4}")
            
    Sheets(shBO).Select
            
    frClB = Sheets(shBO).Cells(Rows.Count, 2).End(xlUp).Row
            
    Sheets(shBO).Range("B" & frClB + 1).Select
    ActiveSheet.Paste
            
    SendKeys "{NUMLOCK}"
        
    frClA = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    frClB = Sheets(shBO).Cells(Rows.Count, 2).End(xlUp).Row
            
    Sheets(shBO).Range("A" & frClA + 1 & " :A" & frClB).Value = nameFile

End Sub

Sub copy_content_up(yearFile As String)
    
    ' copia o conteudo da BO em atualiza_bases
    ' cola na base do ano do arquivo
    
    nameBase = "base_" & yearFile
    
    Workbooks(thisWorkbook.Name).Activate
    Sheets(shBO).Select
    frClA = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(shBO).Range("A2:C" & frClA).Select
    Selection.Copy
    
    Application.Wait (Now + TimeSerial(0, 0, 5))
    
    Workbooks(nameBase).Activate
    Workbooks(nameBase).Sheets(shBO).Select
    frClA = Workbooks(nameBase).Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    Workbooks(nameBase).Sheets(shBO).Range("A" & frClA + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' cola bl
    Workbooks(thisWorkbook.Name).Activate
    Sheets(shBL).Select
    frClA = Sheets(shBL).Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(shBL).Range("A2:T" & frClA).Select
    Selection.Copy
    
    Application.Wait (Now + TimeSerial(0, 0, 5))
    
    Workbooks(nameBase).Activate
    Workbooks(nameBase).Sheets(shBL).Select
    frClA = Workbooks(nameBase).Sheets(shBL).Cells(Rows.Count, 1).End(xlUp).Row
    
    Workbooks(nameBase).Sheets(shBL).Range("A" & frClA + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' cola bc
    Workbooks(thisWorkbook.Name).Activate
    Sheets(shBC).Select
    frClA = Sheets(shBC).Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(shBC).Range("A2:P" & frClA).Select
    Selection.Copy
    
    Application.Wait (Now + TimeSerial(0, 0, 5))
    
    Workbooks(nameBase).Activate
    Workbooks(nameBase).Sheets(shBC).Select
    frClA = Workbooks(nameBase).Sheets(shBC).Cells(Rows.Count, 1).End(xlUp).Row
    
    Workbooks(nameBase).Sheets(shBC).Range("A" & frClA + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Workbooks(nameBase).Save
    
    Application.Wait (Now + TimeSerial(0, 0, 5))
    
    Workbooks(nameBase).Close
    
End Sub




