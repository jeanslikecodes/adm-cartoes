Attribute VB_Name = "md_bas"
Function check_base(yearFile As String)

    Dim baseFile As String

    baseFile = "base_" & yearFile & ".xlsb"
    
    check_base = Dir(basPath & "/" & baseFile)

End Function
 
Sub create_base(yearFile As String)

    ' cria a base (pega modelo em setup) e renomeia o arquivo para o ano que está sendo lido
    
    Dim srcPath As String ' source
    Dim desPath As String ' destiny

    srcPath = CStr(setPath & "\")
    desPath = CStr(basPath & "\")

    srcName = srcPath & "base_modelo.xlsb"
    desName1 = desPath & "base_modelo.xlsb"             ' nome original (base_modelo.xlsb)
    desName2 = desPath & "base_" & yearFile & ".xlsb"   ' nome renomeado (base_year.xlsb)
    
    FileCopy srcName, desName1
    
    Name desName1 As desName2

End Sub

Sub open_base(yearFile As String)

    nameBase = "base_" & yearFile
    
    Workbooks.Open (basPath & "\" & nameBase)
    
    Application.Wait (Now + TimeSerial(0, 0, 4))
    
    Workbooks(nameBase).Activate

End Sub
