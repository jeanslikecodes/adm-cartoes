Attribute VB_Name = "md_exp"
Sub export_code()

    ' exporta os modulos em formato '.bas' _
        p/ que seja lido pelo git e assim _
        fazer o versionamento

    Dim pathCode As String
    Dim thisWB As Workbook
    
    Set thisWB = Workbooks(thisWorkbook.Name)
    
    pathCode = thisWB.path & "/codes"
    
    If Dir(pathCode) <> "" Then
        MkDir (pathCode)
    End If
    
    For Each Module In thisWB.VBProject.VBComponents
        Module.Export (pathCode & "/" & Module.Name & ".bas")
    Next

End Sub
