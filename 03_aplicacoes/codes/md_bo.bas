Attribute VB_Name = "md_bo"
Sub clear_bo()

    Sheets(shBO).Select
    frClA = Sheets(shBO).Cells(Rows.Count, 1).End(xlUp).Row
    
    If frClA > 1 Then
        Sheets(shBO).Range("A2:B" & frClA).Value = ""
    End If

End Sub
