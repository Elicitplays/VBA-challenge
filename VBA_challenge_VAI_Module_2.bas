Attribute VB_Name = "Module1"

Sub test()


    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    Dim opening As Integer
    opening = 2

    Dim first_open As Boolean
    first_open = True
    
    Dim ws As Worksheet
    
    

For Each ws In Worksheets


    Summary_Table_Row = 2
    opening = 2
    first_open = True

        For i = 2 To 22771

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    

                ' Set the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
      
                'Debug.Print (Cells(i + 1, 3).Value)
                'Debug.Print (Cells(i, 6).Value)
                'Debug.Print (Cells(opening, 3).Value)
                'Debug.Print ("___________")
      
                If first_open = True Then
                    opening = 2
                End If
                
                Yearly_change = ws.Cells(i, 6).Value - ws.Cells(opening, 3).Value
      
                Percent_change = (ws.Cells(opening, 3).Value - ws.Cells(i, 6).Value) / (ws.Cells(i, 6).Value)

                ' Add to the volume Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value

                ' Print the Ticker name in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Name

                ' Print the Volume total to the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = Volume_Total
      
                ws.Range("K" & Summary_Table_Row).Value = Yearly_change
      
                ws.Range("L" & Summary_Table_Row).Value = Percent_change
      
      
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
      
                ' Reset the Volume Total
                Volume_Total = 0
      
                first_open = False


            Else

                ' Add to the volume Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
                opening = opening + 1
      
            End If

        Next i
        
        'Debug.Print ("this works once")
        
    Next ws

End Sub


