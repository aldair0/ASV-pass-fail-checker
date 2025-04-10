Attribute VB_Name = "Module2"
Sub CheckFailInColumnG()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellContent As String
    Dim failCount As Long
    Dim passCount As Long
    Dim alterations As String
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row in Column G with data
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Initialize counters and alterations string
    failCount = 0
    passCount = 0
    alterations = "Alterations Made:" & vbNewLine
    
    ' Loop through each row in Column G starting from row 2 (ignoring header)
    For i = 2 To lastRow
        ' Get the content of the current cell in Column G
        cellContent = ws.Cells(i, "G").Value
        
        ' Check if "fail" exists in the cell content (case-insensitive)
        If InStr(1, LCase(cellContent), "fail") > 0 Then
            ' If "fail" is found, set F and N to "Fail" with red text
            With ws.Cells(i, "F")
                .Value = "Fail"
                .Font.Color = vbRed
            End With
            With ws.Cells(i, "N")
                .Value = "Fail"
                .Font.Color = vbRed
            End With
            failCount = failCount + 1
            alterations = alterations & "Row " & i & ": Set to Fail" & vbNewLine
        Else
            ' If "fail" is not found, set F and N to "Pass"
            With ws.Cells(i, "F")
                .Value = "Pass"
            End With
            With ws.Cells(i, "N")
                .Value = "Pass"
            End With
            passCount = passCount + 1
            alterations = alterations & "Row " & i & ": Set to Pass" & vbNewLine
        End If
    Next i
    
    ' Display the message box with totals first, then alterations
    MsgBox "Summary:" & vbNewLine & _
           "Total Fails: " & failCount & vbNewLine & _
           "Total Passes: " & passCount & vbNewLine & vbNewLine & _
           "Processing Complete"
End Sub

