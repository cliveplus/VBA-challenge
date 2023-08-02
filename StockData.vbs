Attribute VB_Name = "Module1"
Sub StockData():
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    TotalVolume = 0
   
    Dim SummaryTable As Integer

 SummaryTable = 2
    
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
  For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        Ticker = ws.Cells(i, 1).Value
        
          ClosingPrice = ws.Cells(i + 1, 6).Value
        
        'format color
            If ws.Range("J" & SummaryTable).Value > 0 Then
             ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
            Else
             ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
            End If
        
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        'Print ticker
        ws.Range("I" & SummaryTable).Value = Ticker
        'Print Total Volume
        ws.Range("L" & SummaryTable).Value = TotalVolume
  
        'Reset Values
        SummaryTable = SummaryTable + 1
        
        TotalVolume = 0
        
        PercentChange = YearlyChange / OpeningPrice
        
        ws.Range("K" & SummaryTable).Value = PercentChange
        ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
        
        Else
        
         OpeningPrice = ws.Cells(i, 3).Value
      
        YearlyChange = ClosingPrice - OpeningPrice
        ws.Range("J" & SummaryTable).Value = YearlyChange
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
       
    End If
      
    Next i
        
        HighestPercent = WorksheetFunction.Max(ws.Columns("K"))
        LowestPercent = WorksheetFunction.Min(ws.Columns("K"))
        HighestVolume = WorksheetFunction.Max(ws.Columns("L"))
        
        ws.Range("O2").Value = HighestPercent
       ws.Range("O3").Value = LowestPercent
       ws.Range("O4").Value = HighestVolume

        
       ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ws.Range("N2").Value = "Greatest Percent Increase"
        ws.Range("N3").Value = "Greatest Percent Decrease"
        ws.Range("N4").Value = "Greatest Volume"
        
        ws.Range("J1").EntireColumn.AutoFit
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("L1").EntireColumn.AutoFit
        
         ws.Range("O2").EntireColumn.AutoFit
         ws.Range("N2").EntireColumn.AutoFit
        
        
        
Next ws
        
        
End Sub


