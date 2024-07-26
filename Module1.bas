Attribute VB_Name = "Module1"
Sub Add_To_Table()
    ' Copy the entire row 5
    Rows(5).Copy

    ' Insert copied row between rows 10 and 11
    Rows(14).Insert Shift:=xlDown

    ' Clear the contents of columns C to S in row 5
    Dim col As Range
    For Each col In Range("C5:S5")
        If col.HasFormula Then
            col.Value = col.Value
        End If
    Next col
End Sub
Sub Save()
    ' Save the workbook
    ThisWorkbook.Save
End Sub
Sub Delete_Top_Row()
    Rows(14).Delete
End Sub
