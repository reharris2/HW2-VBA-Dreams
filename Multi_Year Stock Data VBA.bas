Attribute VB_Name = "Module1"
Sub MultiYear_StockData()

'Script to run the program over multiple sheets
Dim WS As Worksheet
For Each WS In Worksheets
    WS.Activate
    
'Script to run loop that returns the ticker, yearly change, percent change, and total volume
Dim i As Long
Dim Ticker As String

Dim Total_Volume As Double
Total_Volume = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2


Dim Opening_Price As Double
Opening_Price = Cells(2, 3).Value
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double


Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value

Total_Volume = Total_Volume + Cells(i, 7).Value

Range("I" & Summary_Table_Row).Value = Ticker
Range("L" & Summary_Table_Row).Value = Total_Volume

Closing_Price = Cells(i, 6).Value
Yearly_Change = (Closing_Price - Opening_Price)

Range("j" & Summary_Table_Row).Value = Yearly_Change

If Opening_Price = 0 Then
Percent_Change = 0

Else
Percent_Change = (Yearly_Change / Opening_Price)
End If

Range("k" & Summary_Table_Row).Value = Percent_Change

Summary_Table_Row = Summary_Table_Row + 1

Total_Volume = 0
Opening_Price = Cells(i + 1, 3)

Else
Total_Volume = Total_Volume + Cells(i, 7).Value

End If

Next i

Range("I1").Value = "Ticker"
Range("L1").Value = "Total Stock Volume"
Columns("L:L").EntireColumn.AutoFit
Range("J1").Value = "Yearly Change"
Columns("J:J").EntireColumn.AutoFit
Range("K1").Value = "Percent Change"
Columns("K:K").EntireColumn.AutoFit

Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    
    
' Conditional formatting to assign green for positive dollor change and red for negative dollar changes
lastrow_summary_ticker = Cells(Rows.Count, 9).End(xlUp).Row
  
        For i = 2 To lastrow_summary_ticker
            
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.Color = vbGreen
                
            Else
                Cells(i, 10).Interior.Color = vbRed
                
            End If

        Next i
 
'Script to create a table to show the stock with greatest % increase/decrease and greatest total volume
Range("p1").Value = "Ticker"
Range("q1").Value = "Value"
Range("o2").Value = "Greatest % Increase"
Range("o3").Value = "Greatest % Decrease"
Range("o4").Value = "Greatest Total Volume"
Columns("o:o").EntireColumn.AutoFit



lastrow_hard = Cells(Rows.Count, 11).End(xlUp).Row
DMax = 0
For i = 1 To lastrow_hard
    Calculate
    DMax2 = Application.WorksheetFunction.Max(Range("k:k"))
    If DMax2 > DMax Then DMax = DMax2
Next i

Range("q2").Value = DMax


lastrow_hard = Cells(Rows.Count, 11).End(xlUp).Row
DMin = 0
For i = 1 To lastrow_hard
    Calculate
    DMin2 = Application.WorksheetFunction.Min(Range("k:k"))
    If DMin2 < DMin Then DMin = DMin2
Next i

Range("q3").Value = DMin
Range("q2:q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"

lastrow_hard1 = Cells(Rows.Count, 12).End(xlUp).Row
DMax = 0
For i = 1 To lastrow_hard
    Calculate
    DMax2 = Application.WorksheetFunction.Max(Range("l:l"))
    If DMax2 > DMax Then DMax = DMax2
Next i

Range("q4").Value = DMax
Columns("q:q").EntireColumn.AutoFit

Dim percentincrease As Double
Dim percentdecrease As Double
Dim greatvolume As Double

percentincrease = Range("q2").Value
percentdecrease = Range("q3").Value
greatvolume = Range("q4").Value



For i = 2 To lastrow

If Cells(i, 11).Value = percentincrease Then
Cells(2, 16).Value = Cells(i, 9).Value

ElseIf Cells(i, 11).Value = percentdecrease Then
Cells(3, 16).Value = Cells(i, 9).Value

ElseIf Cells(i, 12).Value = greatvolume Then
Cells(4, 16).Value = Cells(i, 9).Value

End If

Next i

Next

 
 
End Sub

  

