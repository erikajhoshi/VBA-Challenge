Attribute VB_Name = "Module1"
Sub Stock()


Dim Ticker As String
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Vol As Double
Vol = 0
Dim Summary_Table_Row As Integer


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Percent_Change"
Cells(1, 12).Value = "Total_Stock_Volume"
Summary_Table_Row = 2


For i = 2 To 70926

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    Vol = Vol + Cells(i, 7)
    
    Year_Open = Cells(i, 3).Value
    Year_Close = Cells(i, 6).Value
    
    Yearly_Change = Year_Close - Year_Open
    Percent_Change = (Year_Close - Year_Open) / Year_Close
    
    
    Range("I" & Summary_Table_Row).Value = Ticker
    Range("L" & Summary_Table_Row).Value = Vol
    Range("J" & Summary_Table_Row).Value = Yearly_Change
    Range("K" & Summary_Table_Row).Value = Percent_Change
    Summary_Table_Row = Summary_Table_Row + 1
    Vol = 0

Else
    Vol = Vol + Cells(i, 7).Value
    Yearly_Change = Year_Close - Year_Open
    
End If

Next i

For i = 2 To 70926

If Cells(i, 10).Value < 0 Then
Cells(i, 10).Interior.ColorIndex = 3

ElseIf Cells(i, 10).Value >= 0 Then
Cells(i, 10).Interior.ColorIndex = 4

End If

Next i

Columns("K").NumberFormat = "0.00%"


End Sub

