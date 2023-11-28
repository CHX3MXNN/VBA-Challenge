Attribute VB_Name = "Module2"
Sub Worksheet_Loop()

 For Each ws In ThisWorkbook.Worksheets
       
       Dim Ticker As String
       Dim YearlyChange As Double
       Dim PercentChange As Double
       Dim TotalVolume As Double
       Dim LastRow As Long
       Dim SummaryRow As Long
       Dim OpeningPrice As Double
       Dim ClosingPrice As Double
       
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
       
       
       LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       
       SummaryRow = 2
       
       
       For i = 2 To LastRow
     
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
               Ticker = ws.Cells(i, 1).Value
               
              
               OpeningPrice = ws.Cells(i, 3).Value
               
               
               ClosingPrice = ws.Cells(i, 6).Value
               
              
               YearlyChange = ClosingPrice - OpeningPrice
               
               
               If OpeningPrice <> 0 Then
                   PercentChange = (YearlyChange / OpeningPrice) * 100
               Else
                   PercentChange = 0
               End If
               
               
               ws.Cells(SummaryRow, 9).Value = Ticker
               ws.Cells(SummaryRow, 10).Value = YearlyChange
               ws.Cells(SummaryRow, 11).Value = PercentChange
               ws.Cells(SummaryRow, 12).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i - TotalVolume + 1, 7), ws.Cells(i, 7)))
               
               
               ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
               
               
               If YearlyChange > 0 Then
                   ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
               ElseIf YearlyChange < 0 Then
                   ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
               End If
               
               
               SummaryRow = SummaryRow + 1
               
               
               TotalVolume = 0
           End If
           
           
           TotalVolume = TotalVolume + ws.Cells(i, 7).Value
       Next i
       
       
       SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
       
       Dim MaxPercentIncrease As Double
       Dim MaxPercent
       
       
Debug.Print myCell.Address(False, False)





End Sub
