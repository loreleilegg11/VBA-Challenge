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
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Quarterly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
            
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest % increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatests Total Volume"
            
            'initialize values
             SummaryTablerow = 0
             TotalVolume = 0
             QuarterlyChange = 0
             StartValue = 2
             StockStartRow = 2
        
        
        Lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        
        'last ticker
        lastTicker = Cells(Lastrow, 1).Value
        
        
        'loop through first column of ticker names
        
        For Row = 2 To Lastrow
        
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
            
            TotalVolume = Cells(Row, 7).Value + TotalVolume
            
            If TotalVolume = 0 Then
                Range("I" & 2 + SummaryTablerow).Value = Cells(Row, 1).Value
                Range("L" & 2 + SummaryTablerow).Value = 0
                Range("J" & 2 + SummaryTablerow).Value = 0
                Range("K" & 2 + SummaryTablerow).Value = 0
            Else
            'first non-zero value for open
                If Cells(StartValue, 3).Value = 0 Then
                    For findValue = StartValue To Row
                    
                    If Cells(findValue, 3).Value <> 0 Then
                        StartValue = findValue
                        Exit For
                End If
                    
                  Next findValue
                
            End If
                 ' calculate change
            QuarterlyChange = Cells(Row, 6).Value - Cells(StartValue, 3).Value
            PercentChange = QuarterlyChange / Cells(StartValue, 3).Value
                 
                 'print results
            
            Range("I" & 2 + SummaryTablerow).Value = Cells(Row, 1).Value
            Range("L" & 2 + SummaryTablerow).Value = TotalVolume
            Range("J" & 2 + SummaryTablerow).Value = QuarterlyChange
            Range("K" & 2 + SummaryTablerow).Value = PercentChange
            
            
            
            'apply colors to change
            Select Case QuarterlyChange
                Case Is > 0
                Range("J" & 2 + SummaryTablerow).Interior.ColorIndex = 4
                Case Is < 0
                Range("J" & 2 + SummaryTablerow).Interior.ColorIndex = 3
                Case Else
                Range("J" & 2 + SummaryTablerow).Interior.ColorIndex = 0
                
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
            
            TotalVolume = Cells(Row, 7).Value + TotalVolume
            
            
         End If
            
     Next Row
            
            SummaryTablerow = Cells(Rows.Count, "I").End(xlUp).Row
            
            'find last data in the extra row from column j-l
        Dim LastExtraRow As Long
        LastExtraRow = Cells(Rows.Count, "J").End(xlUp).Row
            
        For e = SummaryTablerow To LastExtraRow
                For Column = 9 To 12
                    Cells(e, Column).Value = ""
                    Cells(e, Column).Interior.ColorIndex = 0
                Next Column
                
        Next e
            
        ' outside for loop find mins and maxes
            
        Range("Q2").Value = WorksheetFunction.Max(Range("K2:K" & SummaryTablerow + 2))
        Range("Q3").Value = WorksheetFunction.Min(Range("K2:K" & SummaryTablerow + 2))
        Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & SummaryTablerow + 2))
    
        'row numbers of the ticker names for  greates increase and decrease & total value
        
        Dim greatestIncreaseRow As Double
        Dim GreatestDecreaseRow As Double
        Dim GreatesttotalVolRow As Double
        
        greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & SummaryTablerow + 2)), Range("K2:K" & SummaryTablerow + 2), 0)
        GreatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & SummaryTablerow + 2)), Range("K2:K" & SummaryTablerow + 2), 0)
        GreatesttotalVolRow = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & SummaryTablerow + 2)), Range("L2:L" & SummaryTablerow + 2), 0)
        'ticker symbol for values
        
        Range("P2").Value = Cells(greatestIncreaseRow + 1, 9).Value
        Range("P3").Value = Cells(GreatestDecreaseRow + 1, 9).Value
        Range("P4").Value = Cells(GreatesttotalVolRow + 1, 9).Value
        
      For s = 0 To SummaryTablerow
         Range("J" & 2 + SummaryRow).NumberFormat = "0.00"
         Range("K" & 2 + SummaryRow).NumberFormat = "0.00%"
         Range("L" & 2 + SummaryRow).NumberFormat = "#,###"
        Next s
        
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        Range("Q4").NumberFormat = "#,###"
        
        Columns("A:Q").AutoFit
        
   Next ws
   
End Sub


