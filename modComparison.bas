Attribute VB_Name = "Comparison"
Sub CompareAndMark(Cell1 As Range, Cell2 As Range)
    ' Compare values after cleaning each line
    Dim val1 As String
    Dim val2 As String

    val1 = CleanMultiLineText(Cell1.Value)
    val2 = CleanMultiLineText(Cell2.Value)

    If StrComp(val1, val2, vbTextCompare) = 0 Then
        MsgBox "Row " & Cell1.Row & ": Values are Equal", vbInformation, "Comparison Result"
    Else
        MsgBox "Row " & Cell1.Row & ": Values are Not Equal" & vbCrLf & _
                "Value 1: " & val1 & vbCrLf & _
                "Value 2: " & val2, vbExclamation, "Comparison Result"
    End If
End Sub


Function CleanMultiLineText(ByVal strinput As String) As String
    Dim lines As Variant
    Dim line As Variant
    Dim result As String
    Dim cleanLine As String
    
    ' Normalize line endings
    strinput = Replace(strinput, vbLf, vbCrLf)
    strinput = Replace(strinput, vbCr, vbCrLf)
    
    ' Split into lines
    lines = Split(strinput, vbCrLf)
    
    ' Loop through each line
    For Each line In lines
        ' Replace non-breaking space and tabs
        line = Replace(line, Chr(160), " ")
        line = Replace(line, vbTab, " ")
        
        ' Trim only leading/trailing whitespace
        cleanLine = VBA.Trim(line)
        
        ' Skip empty lines
        If cleanLine <> "" Then
            If result = "" Then
                result = cleanLine
            Else
                result = result & vbCrLf & cleanLine
            End If
        End If
    Next line
    
    CleanMultiLineText = result
End Function


Sub TestAllRows_SingleMessage()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Use the currently active worksheet

    Dim i As Long
    Dim LastRow As Long
    Dim msg As String
    msg = "" ' Clear previous messages

    ' Find the last used row in column I or J
    LastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "J").End(xlUp).Row > LastRow Then
        LastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    End If

    For i = 3 To LastRow
        Dim valI As String
        Dim valJ As String

        ' Clean strings: Replace non-breaking space + trim per line
        valI = CleanMultiLineText(ws.Cells(i, "I").Value)
        valJ = CleanMultiLineText(ws.Cells(i, "J").Value)

        ' Compare values case-insensitively
        If StrComp(valI, valJ, vbTextCompare) = 0 Then
            msg = msg & "Row " & i & ": Equal" & vbCrLf

            ' Highlight equal rows in white
            ws.Rows(i).Interior.Color = RGB(255, 255, 255) ' White background
        Else
            msg = msg & "Row " & i & ": Not Equal" & vbCrLf
            
            ' Highlight mismatched rows in light red
            ws.Rows(i).Interior.Color = RGB(255, 204, 204) ' Light red
        End If
    Next i

    ' Show results in UserForm
    With frmComparisonResults
        .txtResults.Text = msg
        .Show vbModeless
    End With
End Sub

