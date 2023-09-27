Attribute VB_Name = "Module2"
Sub ticker_Changes()

    For Each WS In Worksheets
        WS.Activate
       
        'Title Summary Table
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"

        'Set initial variable for holding opener ticker symbol
        open_price = Cells(2, "C").Value
       
       'keep track of the location of ticker symbol in summary table
       Summary_Table_Row = 2
       Start = 2
       total_volume = 0
       
       row_count = Cells(Rows.Count, "A").End(xlUp).Row
       
       'loop through all ticker brands, opening price, and closing price
       For input_row = 2 To row_count
       
        'check if still in same ticker symbol, if not then
        If Cells(input_row + 1, "A").Value <> Cells(input_row, "A").Value Then
       
            total_volume = total_volume + Cells(input_row, "G").Value
       
            'set ticker symbol name
            ticker_name = Cells(input_row, "A").Value
           
           
            If total_volume = 0 Then
                ' print the results
                Range("J" & Summary_Table_Row).Value = Cells(i, "A").Value
                Range("K" & Summary_Table_Row).Value = 0
                Range("L" & Summary_Table_Row).Value = "%" & 0
                Range("M" & Summary_Table_Row).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(Start, "C") = 0 Then
                    For find_value = Start To input_row
                        If Cells(find_value, "C").Value <> 0 Then
                            Start = find_value
                           
                            Exit For
                        End If
                     Next find_value
                End If
               
                 closing_price = Cells(input_row, "F").Value
                 
                 open_price = Cells(Start, "C").Value
           
                yearly_change = closing_price - open_price
               
               
                percentage_change = yearly_change / open_price * 100
               
           
                'print ticker symbol in summary table
                Range("J" & Summary_Table_Row).Value = ticker_name
               
                'print opening price total to summary table
                Range("K" & Summary_Table_Row).Value = yearly_change
               
                'print closing price total to summary table
                Range("L" & Summary_Table_Row).Value = "%" & percentage_change
               
                Range("M" & Summary_Table_Row).Value = total_volume
               
                'add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
               
                'reset the total volume
                total_volume = 0
               
                Start = input_row + 1
               
               
                'reset the open price
                open_price = Cells(input_row + 1, "C")
               
            End If
   
               
                'if the cell immediately following a row is the same brand
        Else
       
            'add to total_volume
            total_volume = total_volume + Cells(input_row, "C").Value
           
       
        End If
       
     Next input_row
   
    Next WS

End Sub
