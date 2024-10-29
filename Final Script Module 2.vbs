Attribute VB_Name = "Module1"
Sub Ticker()
    ' using column A to find when name changes
    ' row 2 to last row
    'declare dimensions
    
    Dim Lastrow As Double
    Dim Row As Long
    Dim TotalVolume As Double
    Dim QuartertlyChange As Double
    Dim PercentChange As Double
    Dim SummaryTablerow As Long
    Dim StockStartRow As Long
    Dim StartValue As Long
    Dim lastTicker As String
   
   For Each ws In Worksheets
    
            'set column headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatests Total Volume"
            
            'initialize values
             SummaryTablerow = 0
             TotalVolume = 0
             QuarterlyChange = 0
             StartValue = 2
             StockStartRow = 2
        
        
        Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'last ticker
        lastTicker = ws.Cells(Lastrow, 1).Value
        
        
        'loop through first column of ticker names
        
        For Row = 2 To Lastrow
        
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
            TotalVolume = ws.Cells(Row, 7).Value + TotalVolume
            
            If TotalVolume = 0 Then
                ws.Range("I" & 2 + SummaryTablerow).Value = ws.Cells(Row, 1).Value
                ws.Range("L" & 2 + SummaryTablerow).Value = 0
                ws.Range("J" & 2 + SummaryTablerow).Value = 0
                ws.Range("K" & 2 + SummaryTablerow).Value = 0
            Else
            'first non-zero value for open
                If ws.Cells(StartValue, 3).Value = 0 Then
                    For findValue = StartValue To Row
                    
                    If ws.Cells(findValue, 3).Value <> 0 Then
                        StartValue = findValue
                        Exit For
                End If
                    
                  Next findValue
                
            End If
                 ' calculate change
            QuarterlyChange = ws.Cells(Row, 6).Value - ws.Cells(StartValue, 3).Value
            PercentChange = QuarterlyChange / ws.Cells(StartValue, 3).Value
                 
                 'print results
            
            ws.Range("I" & 2 + SummaryTablerow).Value = ws.Cells(Row, 1).Value
            ws.Range("L" & 2 + SummaryTablerow).Value = TotalVolume
            ws.Range("J" & 2 + SummaryTablerow).Value = QuarterlyChange
            ws.Range("K" & 2 + SummaryTablerow).Value = PercentChange
            
            
            
            'apply colors to change
            Select Case QuarterlyChange
                Case Is > 0
                ws.Range("J" & 2 + SummaryTablerow).Interior.ColorIndex = 4
                Case Is < 0
                ws.Range("J" & 2 + SummaryTablerow).Interior.ColorIndex = 3
                Case Else
                ws.Range("J" & 2 + SummaryTablerow).Interior.ColorIndex = 0
                
            End Select
            ' reset values
            
            TotalVolume = 0
            SummaryTablerow = SummaryTablerow + 1   ' move to new row
            QuarterlyChange = 0
            PercentChange = 0
            StartValue = Row + 1
            
            
         End If
         
            
         Else
            
            'same ticker add to total
            
            TotalVolume = ws.Cells(Row, 7).Value + TotalVolume
            
            
         End If
            
     Next Row
            
            SummaryTablerow = ws.Cells(Rows.Count, "I").End(xlUp).Row
            
            'find last data in the extra row from column j-l
        Dim LastExtraRow As Long
        LastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
            
        For e = SummaryTablerow To LastExtraRow
                For Column = 9 To 12
                    ws.Cells(e, Column).Value = ""
                    ws.Cells(e, Column).Interior.ColorIndex = 0
                Next Column
                
        Next e
            
        ' outside for loop find mins and maxes
            
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & SummaryTablerow + 2))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & SummaryTablerow + 2))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & SummaryTablerow + 2))
    
        'row numbers of the ticker names for  greates increase and decrease & total value
        
        Dim greatestIncreaseRow As Double
        Dim GreatestDecreaseRow As Double
        Dim GreatesttotalVolRow As Double
        
        greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & SummaryTablerow + 2)), ws.Range("K2:K" & SummaryTablerow + 2), 0)
        GreatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & SummaryTablerow + 2)), ws.Range("K2:K" & SummaryTablerow + 2), 0)
        GreatesttotalVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & SummaryTablerow + 2)), ws.Range("L2:L" & SummaryTablerow + 2), 0)
        'ticker symbol for values
        
        ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(GreatestDecreaseRow + 1, 9).Value
        ws.Range("P4").Value = ws.Cells(GreatesttotalVolRow + 1, 9).Value
        
      For s = 0 To SummaryTablerow
         ws.Range("J" & 2 + SummaryRow).NumberFormat = "0.00"
         ws.Range("K" & 2 + SummaryRow).NumberFormat = "0.00%"
         ws.Range("L" & 2 + SummaryRow).NumberFormat = "#,###"
        Next s
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,###"
        
        ws.Columns("A:Q").AutoFit
        
   Next ws
   
End Sub


