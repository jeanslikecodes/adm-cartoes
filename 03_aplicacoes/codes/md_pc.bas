Attribute VB_Name = "md_pc"
Sub insert_info(nameFile As String, sizefile As Long)

    Sheets(shPC).Select
    frClB = Sheets(shPC).Cells(Rows.Count, 2).End(xlUp).Row
    
    Sheets(shPC).Range("B" & frClB + 1).Value = nameFile
    Sheets(shPC).Range("C" & frClB + 1).Value = CInt(Round(sizefile / 1024, 0)) & " kb"
    Sheets(shPC).Range("D" & frClB + 1).Value = Now()
    
End Sub
