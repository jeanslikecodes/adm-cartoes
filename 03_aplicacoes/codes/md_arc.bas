Attribute VB_Name = "md_arc"
Function check_existence_arq_in_pc(nameFile As String)

    ' checa se o arquivo que irá ser lido consta no painel de contrrole
    
    Dim rwArc As String
    
    Sheets(shPC).Select
    
    On Error Resume Next
    rwArc = WorksheetFunction.Match(nameFile, Sheets(shPC).Range("B:B"), 0)
    
    check_existence_arq_in_pc = rwArc
    
End Function

Sub copy_content_pdf(nameFile As String)

    Dim adobeApp As String
    Dim adobeStart
    
    adobeFile = pdfPath & "/" & nameFile
            
    minutes = 0
    seconds = 45
            
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
    frClA = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(shBO).Range("A2:B" & frClA).Select
    Selection.Copy
    
    Application.Wait (Now + TimeSerial(0, 0, 5))
    
    Workbooks(nameBase).Activate
    Workbooks(nameBase).Sheets(shBO).Select
    frClA = Workbooks(nameBase).Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    Workbooks(nameBase).Sheets(shBO).Range("A" & frClA + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Workbooks(nameBase).Save
    
    Application.Wait (Now + TimeSerial(0, 0, 5))
    
    Workbooks(nameBase).Close
    
End Sub
