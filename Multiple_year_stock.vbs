
Sub stock()

'declare all variable

Dim ws As Worksheet
Dim ticker As String
Dim opens As Double
Dim closes As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim volume As Double
Dim i As Double
Dim lastrow As Double
Dim sumtablerow As Double

'perform the loop for each Worksheet in the workbook

For Each ws In Worksheets
   ws.Cells(1, 9) = "Ticker"
   ws.Range("I1").ColumnWidth = 8
   ws.Cells(1, 10) = "Yearly Change"
   ws.Range("J1").ColumnWidth = 12
   ws.Cells(1, 11) = "Percentage Change"
   ws.Range("K1").ColumnWidth = 16                  'initialise all header inputs and column width
   ws.Cells(1, 12) = "Total Volume"
   ws.Range("L1").ColumnWidth = 16
   ws.Cells(1, 15) = "Ticker"
   ws.Cells(1, 16) = "Value"
   ws.Range("P1").ColumnWidth = 15
   ws.Cells(2, 14) = "Greatest % Increase"
   ws.Range("N1").ColumnWidth = 20
   ws.Cells(3, 14) = "Greatest % Decreased"
   ws.Cells(4, 14) = "Greatest Total Volume"

'initialise all variable to default value
ticker = " "
sumtablerow = 2     'two rows used, lastrow for main data, sumtablerow for result data
opens = 0
closes = 0
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'use the is identify the last row in the worksheet

'starting loop to identify all required values
For i = 2 To lastrow
opens = ws.Cells(sumtablerow, 3).Value  'initiate variable opens with data

'To verify that the existing row and the row after does not have the same value
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 'execute all the following if the condition is met
 ticker = ws.Cells(i, 1).Value
 volume = volume + ws.Cells(i, 7).Value
 closes = ws.Cells(i, 6).Value
 yearlychange = closes - opens
 percentchange = yearlychange / opens   'to get the percentage changed

 ws.Range("I" & sumtablerow).Value = ticker
 ws.Range("J" & sumtablerow).Value = yearlychange
 ws.Range("K" & sumtablerow).Value = percentchange      'to Print out the result data values
 ws.Range("K" & sumtablerow).Style = "Percent"
 ws.Range("L" & sumtablerow).Value = volume
 
 sumtablerow = sumtablerow + 1
 volume = 0
 
 Else
  volume = volume + ws.Cells(i, 7).Value
 
 End If
 
Next i

'declare all variables to find the greatest increase, decrease and total volume
Dim yearchangerow As Double
Dim r As Double
Dim grtstIncrease As Double
Dim grtstDecrease As Double
Dim grtstVolume As Double

'initialise all values
grtstIncrease = 0
grtstDecrease = 0
grtstVolume = 0

yearchangerow = ws.Cells(Rows.Count, 10).End(xlUp).Row 'get the total row of the result data

' starts the loop thru the whole result data
For r = 2 To yearchangerow

' first IF statements to color format the data on the result data
If ws.Cells(r, 10) >= 0 Then
ws.Cells(r, 10).Interior.Color = RGB(109, 255, 109)

Else
ws.Cells(r, 10).Interior.Color = RGB(239, 13, 13) 'when the value is below 0, negative

End If

'to get the Greatest Increase value while looping thru the result data
'as grtstIncrease is initialised with 0, the 1st data trhu the loop will be bigger,
'it will keep looping thru and get the biggest value from the result data
'same goes for variable grtsDecrease, but compaes to the smallest value
If grtstIncrease < ws.Cells(r, 11).Value Then

grtstIncrease = ws.Cells(r, 11).Value

ws.Cells(2, 15).Value = ws.Cells(r, 9).Value
ws.Cells(2, 16).Value = grtstIncrease               'Prints results at another table
ws.Cells(2, 16).Style = "Percent"

ElseIf grtstDecrease > ws.Cells(r, 11).Value Then

grtstDecrease = ws.Cells(r, 11).Value
ws.Cells(3, 15).Value = ws.Cells(r, 9).Value
ws.Cells(3, 16).Value = grtstDecrease               'Prints results at another table
ws.Cells(3, 16).Style = "Percent"

End If

'Another conditional to get Ticker with the grtstVolume,
'similar formula with grtstIncrease as to find the biggest value in the result data
If grtstVolume < ws.Cells(r, 12).Value Then
grtstVolume = ws.Cells(r, 12).Value
ws.Cells(4, 15).Value = ws.Cells(r, 9).Value
ws.Cells(4, 16).Value = grtstVolume                 'Prints results at another table

End If


Next r

      
Next ws
' end all loops, nothing to see here. :D


End Sub




