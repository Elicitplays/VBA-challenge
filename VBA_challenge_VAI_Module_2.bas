Attribute VB_Name = "Module1"


Sub test()
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim opening As Long
    opening = 2
    Dim first_open As Boolean
    first_open = True
    
    Dim ws As Worksheet
    
    
For Each ws In Worksheets
    Summary_Table_Row = 2
    opening = 2
    first_open = True
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                ' Set the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
      
      
                If first_open = True Then
                    opening = 2
                End If
                
                Yearly_change = ws.Cells(i, 6).Value - ws.Cells(opening, 3).Value
      
                Percent_change = (ws.Cells(i, 6).Value - ws.Cells(opening, 3).Value) / (ws.Cells(opening, 3).Value)
                ' Add to the volume Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                ' Print the Ticker name in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
                ' Print the Volume total to the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = Volume_Total
      
                ws.Range("K" & Summary_Table_Row).Value = Yearly_change
                
                If Yearly_change > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbGreen
                ElseIf Yearly_change = 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbWhite
                Else
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbRed
                End If
                
      
                ws.Range("L" & Summary_Table_Row).Value = Percent_change
      
      
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
      
                ' Reset the Volume Total
                Volume_Total = 0
      
                first_open = False
                
                opening = opening + 1
            Else
                ' Add to the volume Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
                opening = opening + 1
      
            End If
        Next i
        
        
        ' this grabs all the values
        ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("l:l"))
        ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("l:l"))
        ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("m:m"))
        
        'formatting and writing the headers
        
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Volume total"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("L:L").NumberFormat = "0.00%"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        
        'Debug.Print ("this works once")
        
    Next ws
End Sub
