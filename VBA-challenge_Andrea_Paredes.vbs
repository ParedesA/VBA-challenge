Attribute VB_Name = "Module1"
Sub VBA_challenge():

For Each ws In Worksheets

'to create headers to columns that will receive info
ws.Range("I1,P1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("Q1") = "Value"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"

'to give a type to the variable
Dim ticker As String
Dim openamount As Double
Dim closeamount As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim totalstock As LongLong
Dim lastrow As LongLong
Dim increase As Double
Dim decrease As Double
Dim greatest As LongLong
Dim Result As String
Dim Result2 As String
Dim Result3 As String
   
totalstock = 0
summaryline = 2
openamount = ws.Range("C2")

'to find last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'to find values for first summary
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            yearlychange = ws.Cells(i, 6) - openamount
            percentchange = yearlychange / openamount
            totalstock = totalstock + ws.Cells(i, 7).Value
            
           ws.Cells(summaryline, 9) = ticker
           ws.Cells(summaryline, 10) = yearlychange
           ws.Cells(summaryline, 11) = percentchange
           ws.Cells(summaryline, 11).NumberFormat = "0.00%"
           ws.Cells(summaryline, 12) = totalstock
           
                'to set conditional formatting column J
                If ws.Cells(summaryline, 10).Value < 0 Then
                    ws.Cells(summaryline, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summaryline, 10).Interior.ColorIndex = 4
                End If
                
           totalstock = 0
           summaryline = summaryline + 1
           openamount = ws.Cells(i + 1, 3)
                      
        Else
           totalstock = totalstock + ws.Cells(i, 7).Value
        End If
    Next i

'to find values to second summary
increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
greatest = Application.WorksheetFunction.Max(ws.Range("L:L"))

ws.Range("Q2") = increase
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3") = decrease
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4") = greatest
ws.Range("Q4").NumberFormat = "0.00E+00"

'to add ticker symbol for second summary
  Result = Application.WorksheetFunction.XLookup(ws.Range("Q2"), ws.Range("K:K"), ws.Range("I:I"), "N/A")
  Result2 = Application.WorksheetFunction.XLookup(ws.Range("Q3"), ws.Range("K:K"), ws.Range("I:I"), "N/A")
  Result3 = Application.WorksheetFunction.XLookup(ws.Range("Q4"), ws.Range("L:L"), ws.Range("I:I"), "N/A")

ws.Range("P2") = Result
ws.Range("P3") = Result2
ws.Range("P4") = Result3

'autofit created columns
ws.Columns("I:Q").AutoFit

Next ws

MsgBox ("Done")

End Sub

