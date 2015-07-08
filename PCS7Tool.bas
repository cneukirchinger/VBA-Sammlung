Attribute VB_Name = "PCS7Tool"
Sub ErzeugeQuelle()
    Dim rRng As Range
    Dim rCell As Range
    Dim rRow As Range
    
    Set rRng = Worksheets("Tabelle1").Range("A2:D33")
    
    For Each rRow In rRng.Rows
        Debug.Print rRow.Cells(1) & " : " & rRow.Cells(2) & " ; //" & rRow.Cells(4)
        For Each rCell In rRow.Cells
        Next rCell
    Next rRow
    Debug.Print " "
    For Each rRow In rRng.Rows
        Debug.Print rRow.Cells(1) & " := " & rRow.Cells(3) & ";"
        For Each rCell In rRow.Cells
        Next rCell
    Next rRow
End Sub
