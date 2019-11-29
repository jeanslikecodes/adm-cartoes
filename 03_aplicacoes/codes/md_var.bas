Attribute VB_Name = "md_var"

' atualiza_bases
Public shPC As String ' painel de controle
Public shBO As String ' base original

Sub instance_variables()
    
    ' centraliza as variaveis mais utilizadas _
        sendo executado ao inicio de um metodo principal _
        facilitando a manutenção e inspenção das variaveis
        
    shPC = "Painel de Controle"
    shBO = "BO"
    
End Sub
