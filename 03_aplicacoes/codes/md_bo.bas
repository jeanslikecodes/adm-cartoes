Attribute VB_Name = "md_bo"
Sub clear_b()

    clear_bo
    clear_bl
    clear_bc

End Sub


Sub clear_bo()

    Sheets(shBO).Select
    frClA = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    If frClA > 1 Then
        Sheets(shBO).Range("A2:C" & frClA).Value = ""
    End If

End Sub

Sub clear_bl()

    Sheets(shBL).Select
    frClA = Sheets(shBL).Cells(Rows.Count, 1).End(xlUp).Row
    
    If frClA > 1 Then
        Sheets(shBL).Range("A2:T" & frClA).Value = ""
    End If

End Sub

Sub clear_bc()

    Sheets(shBC).Select
    frClA = Sheets(shBC).Cells(Rows.Count, 1).End(xlUp).Row
    
    If frClA > 1 Then
        Sheets(shBC).Range("A2:P" & frClA).Value = ""
    End If

End Sub

