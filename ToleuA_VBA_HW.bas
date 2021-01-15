Attribute VB_Name = "Module1"
Sub StockMarket()

    'Loops through all sheets
    For Each ws In Worksheets
    
        'creates variables for opening and closing price of the year
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        
        'creates varibale for holding ticker names
        Dim Ticker_Name As String
        
        'creates variables for yearly and percent change
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        
        'creates variable to hold the total stock volume
        Dim Stock_Total As Double
               Stock_Total = 0
        
        'creates variable to hold the location of each row in summary table and
        'sets the initial row as second row in sheet
        Dim Summary_Table_Row As Double
               Summary_Table_Row = 2
               
        'creates variable to hold the location of the opening price row and
        'sets the initial row as second row in sheet. Many thanks to TA Farshad that helped to understand this part of HW.
        Dim Opening_Price_Row As Double
               Opening_Price_Row = 2
        
        'creates varibale and determined the Last row
        Dim Last_Row As Double
               Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'adds names of the columns to column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'bonus HW
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'used this columns only for testing purposes
        'to check the opening and closing price
        'ws.Cells(1, 13).Value = "Opening_Price"
        'ws.Cells(1, 14).Value = "Closing_Price"
        
        'autofits the cells
        ws.Columns("I:Q").AutoFit
    
            ' loops through all rows
            For i = 2 To Last_Row
            
            'sets the names of current and next row
            Current_Row = ws.Cells(i, 1).Value
            Next_Row = ws.Cells(i + 1, 1).Value
        
                'checks if the tickers in the row are same, if it is not then
                If Current_Row <> Next_Row Then
                
                'sets the Ticker name
                Ticker_Name = Current_Row
                
                'sets the opening and closing price
                Opening_Price = ws.Cells(Opening_Price_Row, 3).Value
                Closing_Price = ws.Cells(i, 6).Value
                
                'calculates the change within a year by subsctracting opening price from closing
                Yearly_Change = Closing_Price - Opening_Price
                
                'calcs the % change within a year. PLNT ticker has 0 opening and closing price,
                'dividing it gives "Overflow", so I set the IF statement to solve that problem
                    If Opening_Price = 0 Then
                       Percent_Change = 0
                    Else: Percent_Change = Yearly_Change / Opening_Price
                    End If
                
                'adds  stock volume to Stock total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                
                'prints all amounts in the Summary tables
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("L" & Summary_Table_Row).Value = Stock_Total
                
                'ws.Range("M" & Summary_Table_Row).Value = Opening_Price
                'ws.Range("N" & Summary_Table_Row).Value = Closing_Price
                
                
                'changes the style of the column to %
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                
                 'increases decimal
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'moves to the next row of the Summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
                'resets the Stock Total
                Stock_Total = 0
                
                'sets the next opening price row
                Opening_Price_Row = i + 1
                
                Else
                
                'adds to the Stock Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
                End If
                
                'performs the conditional formatting that highlights positive and negative change
                If IsEmpty(ws.Cells(i, 10)) Then
                ws.Cells(i, 10).Interior.ColorIndex = 0
                
                ElseIf ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
                Else: ws.Cells(i, 10).Interior.ColorIndex = 3
                
                End If
             
          Next i
          
          'Bonus HW
          
          Max_Percent = WorksheetFunction.Max(ws.Range("K:K"))
          Max_Row_Percent = WorksheetFunction.Match(Max_Percent, ws.Range("K:K"), 0)
          'Debug.Print (Max_Row_Percent)
          ws.Range("Q2") = Max_Percent
          ws.Range("P2") = ws.Cells(Max_Row_Percent, 9).Value
          ws.Range("Q2").Style = "Percent"
          ws.Range("Q2").NumberFormat = "0.00%"
          
          Min_Percent = WorksheetFunction.Min(ws.Range("K:K"))
          Min_Row_Percent = WorksheetFunction.Match(Min_Percent, ws.Range("K:K"), 0)
          'Debug.Print (Max_Row_Percent)
          ws.Range("Q3") = Min_Percent
          ws.Range("P3") = ws.Cells(Min_Row_Percent, 9).Value
          ws.Range("Q3").Style = "Percent"
          ws.Range("Q3").NumberFormat = "0.00%"
          
          Max_Value = WorksheetFunction.Max(ws.Range("L:L"))
          Max_Row_Value = WorksheetFunction.Match(Max_Value, ws.Range("L:L"), 0)
          'Debug.Print (Max_Row_Percent)
          ws.Range("Q4") = Max_Value
          ws.Range("P4") = ws.Cells(Max_Row_Value, 9).Value
          ws.Range("Q4").NumberFormat = "0.0000E+00"
      
          
     Next ws
     

End Sub

