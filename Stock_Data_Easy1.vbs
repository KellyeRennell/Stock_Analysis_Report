Sub Stock_Data_Easy1

Dim Stock_Ticker_Name As String
Dim Total_Stock_Volume as Double
     Total_Stock_Volume = 0
Dim Lastrow As Double
Dim Summary_Table_Row As Double
Summary_Table_Row = 2


   Lastrow = Cells(Rows.Count, "A").End(xlUp).Row


   Cells(1, 9).Value = "Ticker Name"
   Cells(1, 10).Value = "Total Stock Volume"


    ' Loop through stock data volumes
   For i = 2 To LastRow + 1

        ' For when the Stock Ticker Name changes
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
                  ' Set Stock_Ticker_Name
                    Stock_Ticker_Name = Cells(i, 1).Value
                       
                    ' Print the Stock Ticker Name in the Summary Table in Column I
                    Range("I" & Summary_Table_Row).Value = Stock_Ticker_Name
                    ' Print the Total Stock Volume to the Summary Table in Column J
                    Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

                    Summary_Table_Row = Summary_Table_Row + 1
                    'Reset volume
                    Total_Stock_Volume = 0
                    ' If the cell immediately following a row is the same Stock Ticker Name ...

                Else
                        ' Add to Total Stock Volume
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i+1, 7).Value
   End If
 Next i

   Cells(1, 9).Font.Bold = True
  Cells(1, 10).Font.Bold = True
  Cells(1, 9).ColumnWidth = 20
  Cells(1, 10).ColumnWidth = 20

 End Sub
