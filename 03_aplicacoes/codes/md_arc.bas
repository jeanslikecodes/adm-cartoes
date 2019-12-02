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
    seconds = 60
            
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
