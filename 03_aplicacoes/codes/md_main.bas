Attribute VB_Name = "md_main"
Sub include_info_in_base()

    ' ir� fazer a leitura dos arquivos em 01_pdf/ _
        abrir e extrair as informa��es do arquivo em pdf _
        colar na base original temporaria do atualiza_base _
        transformar as informa��es do arquivo e enviar para a 02_base

    md_var.instance_variables
    
    Dim nameFile As String
    Dim yearFile As String
     
    For Each myFile In pdfPath.Files
        
        nameFile = myFile.Name
        
        If Right(nameFile, 4) = ".pdf" Or _
            Left(nameFile, 1) = "~" Then
            
            ' verifica se o arquivo j� foi lido
            If md_arc.check_existence_arq_in_pc(nameFile) = "" Then
                
                yearFile = Left(nameFile, 4)
                
                ' verificar se a base do ano do arquivo existe
                If md_bas.check_base(yearFile) = "" Then
                    ' cria base em 02_base
                    md_bas.create_base yearFile
                End If
                
                ' copiar conteudo do pdf pro arquivo
                md_arc.copy_content_pdf nameFile
                
                ' formatar arquivo
                
                ' abrir base
                md_bas.open_base yearFile
                
                ' copiar p/ base
                md_arc.copy_content_up yearFile
                
                ' atualizar painel de controle
                
            End If
            
        End If
    
    Next

End Sub


