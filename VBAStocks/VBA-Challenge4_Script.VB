Sub stockFinder()

  ' Set an initial variable for holding the ticker name
  Dim ticker As String

  ' Set an initial variable for holding the total ticker volume
  Dim tickerTotal As Double
  tickerVolumeTotal = 0
  
  ' Set ticker math functions & date
  Dim percentChange As Long
  percentChange = 0
  
  Dim priceChange As Double
  priceChange = 0
  
  Dim tickerDate As Date 'Not sure if should set as date or long since format is
  tickerDate = 0
  
  ' Set ticker parameters for math
  Dim tickerOpen As Double
  tickerOpen = 0
  
  Dim tickerHigh As Double
  tickerHigh = 0
  
  Dim tickerLow As Double
  tickerLow = 0
  
  Dim tickerClose As Double
  tickerClose = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim SummaryTableRow As Integer
  SummaryTableRow = 2
  
  'This will loop through each tab of the workbook.
  Dim ws As Worksheet
  
'****************************************************************


' Add headers for summary tables

Range("J1").Value = "Ticker"
Range("K1").Value = "Total Volume"
Range("L1").Value = "Percent % Change"
Range("M1").Value = "Annual  $ Change"

Range("O1").Value = "Greatest % Increase"
Range("O2").Value = "Greatest % Decrease"
Range("O3").Value = "Greatest Total Volume"
  
  
  '****************************************************************

'For Each ws In Worksheets
    'able to print on multiple tabs, but not able to get this to process on all sheets.
    

  ' Loop through all the different tickers
  For I = 2 To 100000

    ' Check if still looking at same ticker name, when not equal to previous any longer
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the column locations for parameters
      ticker = Cells(I, 1).Value
      tickerOpen = Cells(I, 3).Value
      tickerClose = Cells(I, 6).Value
      tickerHigh = Cells(I, 4).Value
      tickerLow = Cells(I, 5).Value
      

      ' Add all the ticker volume for  the entire days of ticker measured
      tickerVolumeTotal = tickerVolumeTotal + Cells(I, 7).Value
      
      ' Add the percent change = ((close - open) / closed)) *100
      'percentChange = tickerClose - tickerOpen / tickerOpen
      percentChange = Cells(I, 3).Value - Cells(I, 6).Value ' this is not working
      
      ' Add the yearly price change BOR vs EOR per ticker
      ' tickerHigh - tickerLow
      priceCompare = tickerHigh - tickerLow 'for a given date range :/
      
      '****************************************************************
      
      ' Print the individual ticker names in summary table
      Range("J" & SummaryTableRow).Value = ticker

      ' Print the total ticker volume of stocks sold
      Range("K" & SummaryTableRow).Value = tickerVolumeTotal
      
      ' Print the percent change related to ticker name
      Range("L" & SummaryTableRow).Value = percentChange
      
      ' Print the price change for ticker BOR vs EOR
      Range("M" & SummaryTableRow).Value = priceCompare
      
      '****************************************************************

      ' Add one to the summary table row
      SummaryTableRow = SummaryTableRow + 1
      
      ' Reset the ticker volume total
      tickerVolumeTotal = 0

    ' If proceeding cell is same ticker
    Else

      ' Add to the total volume of share sold for individual ticker
      tickerVolumeTotal = tickerVolumeTotal + Cells(I, 7).Value


    End If

  Next I
  
'ws.Activate

'This will loop through multiple sheets in a workbook



End Sub

Sub conditionalFormat() '(ByVal Target As Range)

    Set I = Intersect(Target, Range("K2:K300"))
    If Not I Is Nothing Then
    
        Select Case Target
            Case 0 To 1000000: NewColor = 37 'light blue
            Case 101 To 2000000: NewColor = 46 'orange
            Case 201 To 3000000: NewColor = 12 'dark yellow
            Case 301 To 4000000: NewColor = 10 'green
            Case 401 To 10000000: NewColor = 3 'red
            
        End Select
        
        Target.Interior.ColorIndex = NewColor
    End If

End Sub

Sub ClearCells()
Range("J1:O500").ClearContents

End Sub


Sub sheetLoop()

'This will loop through multiple sheets in a workbook

Dim ws As Worksheet
For Each ws In Worksheets

    ws.Cells(5, 8) = "This works!"

    ws.Activate
    Debug.Print ws.Name
Next

    
End Sub



