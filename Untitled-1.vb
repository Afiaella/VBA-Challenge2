Sub Stock_Market():

'Define the variables 
Dim ticker As String 
Dim yearly_open As Double
 Dim high As Double 
 Dim low As Double 
 Dim yearly_close As Double 
 Dim vol As Double 
 Dim yearly_change As Double
 Dim percentage_change As Double 
 Dim Total_stock_volume As Long 
 Dim i As Long

'Assign a value to the variables 
Cells(1, 1).Value = "ticker" 
Cells(1, 2).Value = "date" 
Cells(1, 3).Value = "yearly_open" 
Cells(1, 4).Value = "high" 
Cells(1, 5).Value = "low" 
Cells(1, 6).Value = "yearly_close" 
Cells(1, 7).Value = "vol"

'Create new table and assign value to variables 
Cells(1, 10).Value = "ticker" 
Cells(1, 11).Value = "yearly_change" 
Cells(1, 12).Value = "percentage_change" 
Cells(1, 13).Value = "total_stock_volume"

Summary_Table_Row = 2

'Loop 
For i = 2 To RowCount 
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ticker = Cells(i,1).value vol = Cells(i,7).value

yearly_open = Cells(i,3).value yearly_close = Cells(i,6).value

yearly_change = yearly_close - yearly_open percentage_change = (yearly_close - yearly_open) / yearly_close

Cells(Summary_Table_Row, 10).Value = ticker 
Cells(Summary_Table_Row, 11).Value = yearly_change 
Cells(Summary_Table_Row, 12).Value = percentage_change 
Cells(Summary_Table_Row, 13).Value = Total_stock_volume Summary_Table_Row = Summary_Table_Row + 1

vol = 0

End If

Next i

End Sub