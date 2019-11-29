Attribute VB_Name = "md_main"
Sub include_info_in_base()

    ' irá fazer a leitura dos arquivos em 01_pdf/ _
        abrir e extrair as informações do arquivo em pdf _
        colar na base original temporaria do atualiza_base _
        transformar as informações do arquivo e enviar para a 02_base

    md_var.instance_variables
    
    Dim file As String
     
    For Each myFile In pdfPath.Files
        
        nameFile = myFile.Name
        
        If Right(nameFile, 4) = ".pdf" Or _
            Left(nameFile, 1) = "~" Then
            
            md_arc.check_existence nameFile
            
        End If
    
    Next

End Sub

