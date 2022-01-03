Attribute VB_Name = "Module2"
Sub Multiple_year_stock_data() 'to group stocks, they were labeled them with letters in format (A, AA, AAB e.t.c.)


'Print column headers
  
    cells(1, 10).Value = "Ticker"
    cells(1, 11).Value = "Yearly Change"
    cells(1, 12).Value = "Percent Change"
    cells(1, 13).Value = "Total Stock Volume"
 
'Set an initial variables for holding ticker, total volume,
  
    Dim Ticker As String
    Dim Total_Volume As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
'Set initial values
    Ticker = cells(2, 1).Value
    Total_Volume = 0
    Opening_Price = cells(2, 3).Value
    Closing_Price = 0
    Percent_Change = 0
    

'Keep track of the location for each stock in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    lastRow = Sheet2.cells(Rows.Count, "A").End(xlUp).Row + 1

    
    For i = 2 To lastRow
    
    
        If cells(i, 1) = Ticker Then

'Continuing same ticker
            Total_Volume = Total_Volume + cells(i, 7).Value
            Closing_Price = cells(i, 6).Value
        
        Else
' Starting  new ticker.

' First, let us finish with previous ticker
            Yearly_Change = Closing_Price - Opening_Price
            Percent_Change = Yearly_Change / Opening_Price
            cells(Summary_Table_Row, 10).Value = Ticker
            cells(Summary_Table_Row, 11).Value = Yearly_Change
            cells(Summary_Table_Row, 12).Value = Percent_Change
            cells(Summary_Table_Row, 13).Value = Total_Volume
        
' Now, let us start with the new ticker
            Summary_Table_Row = Summary_Table_Row + 1
            Ticker = cells(i, 1).Value
            Total_Volume = cells(i, 7).Value
            Opening_Price = cells(i, 3).Value
            Closing_Price = cells(i, 6).Value
        
        End If
        
    Next i

End Sub













