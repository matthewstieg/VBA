Attribute VB_Name = "Module1"
Sub Stock()

For Each ws In Worksheets


       Dim Ticker As String
       Dim TotalVolume As Double
       Dim SummaryRow As Double
       
       lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       SummaryRow = 2
       
            ws.Range("I1") = "Ticker"
            ws.Range("L1") = "Total Stock Volume"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
       
       
            For i = 2 To lastrow
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Ticker = ws.Cells(i, 1).Value
                    TotalVolume = TotalVolume + ws.Cells(i, 7)
                    ws.Range("I" & SummaryRow).Value = Ticker
                    ws.Range("L" & SummaryRow).Value = TotalVolume
                    SummaryRow = SummaryRow + 1
                    TotalVolume = 0
                
                Else
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    
                End If
                
           Next i

                Dim YearOpen As Double
                Dim YearClose As Double
                Dim YearChange As Double
                Dim PercentChange As Double
                
                
            SummaryRow = 2
            
           For i = 2 To lastrow
                
            
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    YearOpen = ws.Cells(i, 3).Value
                
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    YearClose = ws.Cells(i, 6)
                    YearChange = YearClose - YearOpen
                    ws.Range("J" & SummaryRow).Value = YearChange
                        If YearChange >= 0 Then
                        ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                        Else
                        ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                        End If
                    If YearOpen = 0 Then
                    PercentChange = 0
                    Else: PercentChange = (YearChange / YearOpen)
                    End If
                    ws.Range("K" & SummaryRow).Value = PercentChange
                    ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
                    SummaryRow = SummaryRow + 1
                    YearChange = 0
                End If
                    
            Next i

Next ws
            
End Sub
