Attribute VB_Name = "PCS7Tool"
Sub ErzeugeQuelle()
    Dim rRng As Range
    Dim rCell As Range
    Dim rRow As Range
    Dim w As Worksheet
    Dim iReserve As Integer
    Dim sSource As String
    
    For Each w In Worksheets
        sSource = sSource & GenerateHeader(w.Name, "", "ZETA", "Roche", "AS_KOMM", "1.0")
        
        Set rRng = w.Range("A2:D33")
        
        'STRUCT
        sSource = sSource & "STRUCT" & Chr(10)
        sSource = sSource & Chr(9) & "Watchdog : INT ; //Kommunikationsüberwachung" & Chr(10)
        
        For Each rRow In rRng.Rows
            sSource = sSource & Chr(9) & rRow.Cells(1) & " : " & rRow.Cells(2) & " ; //" & rRow.Cells(4) & Chr(10)
            For Each rCell In rRow.Cells
            Next rCell
        Next rRow
        sSource = sSource & Chr(9) & "Reserve : ARRAY  [2 .. 238 ] OF BYTE ;" & Chr(10)
        sSource = sSource & "END_STRUCT ;" & Chr(10)
        
        'DATA
        sSource = sSource & "BEGIN" & Chr(10)
        sSource = sSource & Chr(9) & "Watchdog := 0;" & Chr(10)
        For Each rRow In rRng.Rows
            sSource = sSource & Chr(9) & rRow.Cells(1) & " := " & rRow.Cells(3) & ";" & Chr(10)
            For Each rCell In rRow.Cells
            Next rCell
        Next rRow
        For iReserve = 2 To 238
            sSource = sSource & Chr(9) & "Reserve[" & CStr(iReserve) & "] := B" & Chr(35) & "16" & Chr(35) & "0;" & Chr(10)
        Next iReserve
        sSource = sSource & "END_DATA_BLOCK" & Chr(10) & Chr(10)
    Next w
    
    WriteTextFile (sSource)
End Sub

Private Function GenerateHeader(sDataBlock As String, sTitle As String, sAuthor As String, sFamily As String, sName As String, sVersion As String)
    Dim s As String
    
    s = "DATA_BLOCK " & Chr(34) & sDataBlock & Chr(34) & Chr(10)
    s = s & "TITLE = " & sTitle & Chr(10)
    s = s & "AUTHOR : " & sAuthor & Chr(10)
    s = s & "FAMILY : " & sFamily & Chr(10)
    s = s & "NAME : " & sName & Chr(10)
    s = s & "Version : " & sVersion & Chr(10)
    
    GenerateHeader = s

End Function

Private Sub ChrTable()
    Dim i As Integer
    Dim s As String
    
    For i = 1 To 127
        s = s & CStr(i) & ": " & Chr(i) & Chr(10)
    Next i
    WriteTextFile (s)
End Sub

Private Function WriteTextFile(sContent As String, Optional sFilePath As String = "C:\temp\test.txt")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(sFilePath)
    
    oFile.Write sContent
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing

End Function
