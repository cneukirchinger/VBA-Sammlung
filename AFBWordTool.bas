Attribute VB_Name = "AFBWordTool"
Sub AuslesenWord()
    Dim WordApp As New Word.Application
    Dim t As Table
    Dim c As Cell
    
    Dim s As String
    
    

    With WordApp
        .Visible = True
        .Documents.Open Filename:="C:\Users\chn\Downloads\AFBs\SIP_B1010_F14X6\661_AFB_4B1420_SIP_B1010_F14X6x_CHOB4_V0.7.doc"
            
    For Each t In .ActiveDocument.Range.Tables
        s = Trim(Left(t.Cell(1, 1).Range.Text, Len(t.Cell(1, 1).Range.Text) - 1))
        If InStr(1, s, "EQM (Unit)") > 0 Then
            For Each c In t.Range.Cells
                s = c.Range.Text
                s = Left(s, Len(s) - 1)
                If c.ColumnIndex = 1 And c.RowIndex <> 1 Then
                    Debug.Print s
                    Dim eqm, unit As String
                    eqm = Left(s, InStr(s, "(") - 2)
                    Debug.Print eqm
                    unit = Right(s, InStr(s, "(") - 1)
                    unit = Left(unit, Len(unit) - 2)
                    Debug.Print unit
                End If
            Next c
        Else
        End If
    Next t
    

    End With
    
    WordApp.Quit
    
End Sub

