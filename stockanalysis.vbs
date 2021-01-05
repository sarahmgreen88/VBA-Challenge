Sub stockanalysis()
Dim WS As Worksheet
  'Create header titles: Ticker, Yearly Change, Percent Change and Total Stock Volume
For Each WS In Worksheets
            WS.Cells(1, 9).Value = "Ticker"
            WS.Cells(1, 10).Value = "Yearly Change"
            WS.Cells(1, 11).Value = "Percent Change"
            WS.Cells(1, 12).Value = "Total Stock Volume"
            'Create headers for the bonus
            WS.Cells(2, 15).Value = "Greatest % Increase"
            WS.Cells(3, 15).Value = "Greatest % Decrease"
            WS.Cells(4, 15).Value = "Greatest Total Volume"
            WS.Cells(1, 16).Value = "Ticker"
            WS.Cells(1, 17).Value = "Value"
            
            
        Dim lastrow As Long
    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
   ' Set an initial variable for holding the brand name
  Dim Ticker As String
  Ticker = WS.Cells(2, 1).Value
  ' Set an initial variable for holding the total per credit card brand
  Dim YearlyChange As Double
  'Set an initial variable for percent change
  Dim PercentChange As Double
  'Set an initial variable for Total Stock Volume
  Dim TotalStockVolume As Double
  TotatlStockVolume = 0
  'Create variables for year open and year close
  Dim YearOpen As Double
  Dim YearClose As Double
  Dim g As Double
  g = 2
  ' Keep track in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  ' Loop through all Tickers
    For i = 2 To lastrow

    ' Check if the name of the ticker is different, if it is not...
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker = WS.Cells(i + 1, 1).Value
      
      
      'Create Variables for year open and year close
      YearOpen = WS.Cells(g, 3).Value
      YearClose = WS.Cells(i, 6).Value
      'Determine YearlyChange
      YearlyChange = YearClose - YearOpen
      'Add to Percent Change
      If YearOpen = 0 Then
      PercentChange = 0
      Else
        PercentChange = YearlyChange / YearOpen
        End If
       
      'Add to the TotalStock Volume
      TotalStockVolume = TotalStockVolume + WS.Cells(i, 7).Value
      
      
      ' Print the Ticker Name in the Summary Table
     WS.Cells(Summary_Table_Row, 9).Value = WS.Cells(i, 1).Value
      ' Print the YearlyChange to the Summary Table
      WS.Cells(Summary_Table_Row, 10).Value = YearlyChange
      'Print the PercentChange to the Summary Table
      WS.Cells(Summary_Table_Row, 11).Value = PercentChange
      'Assign column value to TotalStockVolume
      WS.Cells(Summary_Table_Row, 12).Value = TotalStockVolume
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      If YearlyChange > 0 Then
      WS.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
      Else
      WS.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
      End If
      WS.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
      ' Reset the TotalStockVolume
     TotalStockVolume = 0
    ' Track opening value of next ticker
    g = i + 1
    Else
      ' Add to the TotalStockVolume
      TotalStockVolume = TotalStockVolume + WS.Cells(i, 7).Value
    End If
  Next i
  WS.Cells(2, 17).Value = WorksheetFunction.Max(WS.Range("K:K"))
  WS.Cells(3, 17).Value = WorksheetFunction.Min(WS.Range("K:K"))
  WS.Cells(4, 17).Value = WorksheetFunction.Max(WS.Range("L:L"))
  For j = 2 To lastrow
  If WS.Cells(j, 11).Value = WS.Cells(2, 17).Value Then
            WS.Cells(2, 16).Value = WS.Cells(j, 9).Value
    ElseIf WS.Cells(j, 11).Value = WS.Cells(3, 17).Value Then
            WS.Cells(3, 16).Value = WS.Cells(j, 9).Value
    ElseIf WS.Cells(j, 12).Value = WS.Cells(4, 17).Value Then
            WS.Cells(4, 16).Value = WS.Cells(j, 9).Value
            End If
    Next j
    WS.Cells(2, 17).NumberFormat = "0.00%"
     WS.Cells(3, 17).NumberFormat = "0.00%"
     WS.Cells(4, 17).NumberFormat = "0"
  Next
End Sub

