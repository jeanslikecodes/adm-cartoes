Attribute VB_Name = "md_arc"
Function check_existence_arq_in_pc(nameFile As String)

    ' checa se o arquivo que irá ser lido consta no painel de contrrole
    
    Dim rwArc As String
    
    Sheets(shPC).Select
    
    On Error Resume Next
    rwArc = WorksheetFunction.Match(nameFile, Sheets(shPC).Range("B:B"), 0)
    
    check_existence_arq_in_pc = rwArc
    
End Function
