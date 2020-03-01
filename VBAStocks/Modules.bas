Attribute VB_Name = "Module1"



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


Function getAllstockdata() As Object

 

 Dim totalData
Set totalData = CreateObject("Scripting.Dictionary")
 
 tradeDateOpen = 22220101
 tradeDateClose = 0
 Dim editBucket As stockdata
 
Set allTheRows = ActiveSheet
For Each rw In allTheRows.Rows
  If allTheRows.Cells(rw.Row, 1).Value = "" Then
    Exit For
  End If
 If (totalData.Exists(allTheRows.Cells(rw.Row, 1)) = False) Then
 Dim retVal As stockdata
 retVal.ticker = inputticker
 retVal.percentchange = 0
 retVal.totalstockvolume = 0
 retVal.yearchange = 0
 editBucket = retVal
  totalData.Add allTheRows.Cells(rw.Row, 1), retVal
  Else
  Set editBucket = totalData.Item(allTheRows.Cells(rw.Row, 1))
 End If
  If (allTheRows.Cells(rw.Row, 2) < editBucket.tradeDateOpen) Then
   editBucket.tradeDateOpen = allTheRows.Cells(rw.Row, 2)
   editBucket.yearOpen = allTheRows.Cells(rw.Row, 3)
   End If
   If (allTheRows.Cells(rw.Row, 2) > editBucket.tradeDateClose) Then
   editBucket.tradeDateClose = allTheRows.Cells(rw.Row, 2)
   editBucket.yearClose = allTheRows.Cells(rw.Row, 3)
   End If
   editBucket.totalstockvolume = editBucket.totalstockvolume + allTheRows.Cells(rw.Row, 7)


 Next rw
 a = totalData.Items             'Get the items
For i = 0 To totalData.Count - 1 'Iterate the array
 Dim bucket As stockdata
 Set bucket = a(i)
 Debug.Print bucket.ticker
 bucket.yearchange = bucket.yearClose - bucket.yearOpen
 bucket.percentchange = (bucket.yearClose - bucket.yearOpen) / (bucket.yearOpen) * 100
Next
 
 getAllstockdata = totalData

End Function

Sub testalldata()
    getAllstockdata
End Sub

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





