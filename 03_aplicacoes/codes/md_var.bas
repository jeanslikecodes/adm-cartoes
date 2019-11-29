Attribute VB_Name = "md_var"

' atualiza_bases
Public shPC As String ' painel de controle
Public shBO As String ' base original

Public path00 As String
Public path01 As String
Public path02 As String
Public path03 As String

Public setPath As Object
Public pdfPath As Object
Public basPath As Object
Public appPath As Object


Sub instance_variables()
    
    ' centraliza as variaveis mais utilizadas _
        sendo executado ao inicio de um metodo principal _
        facilitando a manutenção e inspenção das variaveis
        
    shPC = "Painel de Controle"
    shBO = "BO"
    
    pathVar = Split(thisWorkbook.path, "\")
    pathLen = Len(pathV(UBound(pathV)))
    path03 = thisWorkbook.path
    path02 = Left(path03, Len(path03) - pathLen) & "02_base"
    path01 = Left(path03, Len(path03) - pathLen) & "01_arquivos"
    path00 = Left(path03, Len(path03) - pathLen) & "00_setup"
    
    Set obj = New Scripting.FileSystemObject
    
    Set setPath = obj.GetFolder(path00)
    Set pdfPath = obj.GetFolder(path01)
    Set basPath = obj.GetFolder(path02)
    Set appPath = obj.GetFolder(path03)
    

End Sub
