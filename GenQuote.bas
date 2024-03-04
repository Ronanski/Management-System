Attribute VB_Name = "GenQuote"
Sub LastStep()
    
    Dim wb As Workbook
    Dim wr As Worksheet
    Dim wf As Worksheet
    Dim colNum As Integer
    Dim qNum As Integer
    
    Set wb = ActiveWorkbook
    Set wr = wb.Sheets("Registry")
    Set wf = wb.Sheets("Forms")
    
    ' Initialize
    wf.Range("F1:N100").Select
    With Selection
    .ClearContents
    .NumberFormat = "General"
    End With
    
    
    ' Array for Company Format
    quoteSet = Array("Group Title", "Item Description", "Lead Time", "Qty", "Units", _
            "Initial Cost", "Mark Up", "Unit Cost", "Final Amount")
    quoteFormat = wr.Range("F27").Value
    quoteType = wr.Range("F15").Value
    currSelect = wr.Range("F25").Value
                    
    ' For Company Format
    If quoteFormat = "Company Format" Then
        ' if Product
        If quoteType = "Product" Then
            For qNum = 1 To 9
                    wf.Cells(1, qNum + 5).Value = quoteSet(qNum - 1)
            Next qNum
        ' if Services
        ElseIf quoteType = "Services" Then
            For qNum = 1 To 9
                If qNum = 3 Then
                    qNum = 4
                End If
                wf.Cells(1, qNum + 5).Value = quoteSet(qNum - 1)
            Next qNum
            ' Delete column as Lead time is not included
            wf.Range("H:H").Delete
        End If
    Else
        ' For Customer Format
        colNum = 6
        For rowNum = 1 To 10
            If wr.Cells(rowNum + 13, 8).Value = "True" Then
                wf.Cells(1, colNum).Value = wr.Cells(rowNum + 13, 7).Value
                colNum = colNum + 1
            End If
        Next rowNum
    End If
    
    ' Fomatting the Cells
    wf.Range("F1", Range("O1").End(xlToLeft)).Select
    With Selection
    .HorizontalAlignment = xlCenter
    .Columns.ColumnWidth = 18
    .Font.Bold = True
    End With
    
    ' UNIT COST, LABOR COST, FINAL AMOUNT FORMAT MUST BE "$" AND "     #,##0.00"
    ' MARK UP FORMAT MUST BE "#0,00"
    
    Select Case currSelect
        Case "Peso"
            numFormat = ChrW(8369) & "     #,##0.00" ' t*ng 'nang Peso Sign 'to
        Case "Dollar"
            numFormat = "$     #,##0.00"
        Case "Euro"
            numFormat = "[$£-en-GB]     #,##0.00"
        Case "Yen"
            numFormat = "[$¥-ja-JP]     #,##0.00"
    End Select
    
    For colNum = 6 To 14
    colSelect = wf.Cells(1, colNum).Value
    
    Select Case colSelect
        Case "Initial Cost"
            wf.Range(Cells(1, colNum), Cells(100, colNum)).Offset(1, 0).Select
            Selection.NumberFormat = numFormat
        Case "Unit Cost"
            wf.Range(Cells(1, colNum), Cells(100, colNum)).Offset(1, 0).Select
            Selection.NumberFormat = numFormat
        Case "Labor Cost"
            wf.Range(Cells(1, colNum), Cells(100, colNum)).Offset(1, 0).Select
            Selection.NumberFormat = numFormat
        Case "Final Amount"
            wf.Range(Cells(1, colNum), Cells(100, colNum)).Offset(1, 0).Select
            Selection.NumberFormat = numFormat
        Case "Mark Up"
            wf.Range(Cells(1, colNum), Cells(100, colNum)).Offset(1, 0).Select
            Selection.NumberFormat = "#0.00%"
    End Select
    
    Next colNum
    
End Sub

