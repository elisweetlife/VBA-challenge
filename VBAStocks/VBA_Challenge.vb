Type stockdata

ticker As String
yearchange As Double
percentchange As Double
totalstockvolume As Double
End Type


Function getstockdata(inputticker As String) As stockdata

 Dim retVal As stockdata
 Dim tradeDateOpen As Long
 Dim tradeDateClose As Long
 Dim yearOpen As Double
 Dim yearClose As Double
 
 tradeDateOpen = 22220101
 tradeDateClose = 0
 retVal.ticker = inputticker
 retVal.percentchange = 0
 retVal.totalstockvolume = 0
 retVal.yearchange = 0
 
Set allTheRows = ActiveSheet
For Each rw In allTheRows.Rows
  If allTheRows.Cells(rw.Row, 1).Value = "" Then
    Exit For
  End If
 If (StrComp(inputticker, allTheRows.Cells(rw.Row, 1), vbTextCompare) = 0) Then
  If (allTheRows.Cells(rw.Row, 2) < tradeDateOpen) Then
   tradeDateOpen = allTheRows.Cells(rw.Row, 2)
   yearOpen = allTheRows.Cells(rw.Row, 3)
   End If
   If (allTheRows.Cells(rw.Row, 2) > tradeDateClose) Then
   tradeDateClose = allTheRows.Cells(rw.Row, 2)
   yearClose = allTheRows.Cells(rw.Row, 3)
   End If
   retVal.totalstockvolume = retVal.totalstockvolume + allTheRows.Cells(rw.Row, 7)
   Debug.Print yearClose
 End If
 
 Next rw
 
 retVal.yearchange = yearClose - yearOpen
 retVal.percentchange = (yearClose - yearOpen) / (yearOpen) * 100
 
 getstockdata = retVal

End Function


Function getstockdata(inputticker As String) As stockdata

 Dim retVal As stockdata
 Dim tradeDateOpen As Long
 Dim tradeDateClose As Long
 Dim yearOpen As Double
 Dim yearClose As Double
 
 tradeDateOpen = 22220101
 tradeDateClose = 0
 retVal.ticker = inputticker
 retVal.percentchange = 0
 retVal.totalstockvolume = 0
 retVal.yearchange = 0
 
Set allTheRows = ActiveSheet
For Each rw In allTheRows.Rows
  If allTheRows.Cells(rw.Row, 1).Value = "" Then
    Exit For
  End If
 If (StrComp(inputticker, allTheRows.Cells(rw.Row, 1), vbTextCompare) = 0) Then
  If (allTheRows.Cells(rw.Row, 2) < tradeDateOpen) Then
   tradeDateOpen = allTheRows.Cells(rw.Row, 2)
   yearOpen = allTheRows.Cells(rw.Row, 3)
   End If
   If (allTheRows.Cells(rw.Row, 2) > tradeDateClose) Then
   tradeDateClose = allTheRows.Cells(rw.Row, 2)
   yearClose = allTheRows.Cells(rw.Row, 3)
   End If
   retVal.totalstockvolume = retVal.totalstockvolume + allTheRows.Cells(rw.Row, 7)
   Debug.Print yearClose
 End If
 
 Next rw
 
 retVal.yearchange = yearClose - yearOpen
 retVal.percentchange = (yearClose - yearOpen) / (yearOpen) * 100
 
 getstockdata = retVal

End Function



Sub TestYearChange()
    Dim testReturn As stockdata
    testReturn = getstockdata("A")
    MsgBox testReturn.yearchange, vbYesNo
End Sub

Sub TestTotalStockVolume()
    Dim testReturn As stockdata
    testReturn = getstockdata("A")
    MsgBox testReturn.totalstockvolume, vbYesNo
End Sub


Sub TestPercentageChange()
    Dim testReturn As stockdata
    testReturn = getstockdata("A")
    MsgBox testReturn.percentchange, vbYesNo
End Sub















