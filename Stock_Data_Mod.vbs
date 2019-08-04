Sub Stock_Data_Mod()
    Dim Sheet as Worksheet
    Dim Stock_Ticker_Name As String 
    Dim Total_Stock_Volume as Double
      Total_Stock_Volume = 0
    Dim Summary_Table_Row As Double
     
    Dim Annual_Change As Double
    Dim Annual_Percent_Change As Double 
    Dim Annual_Open_Price As Integer
    Dim Annual_Close_Price As Integer 
    Dim Lastrow as Long   
    
    Lastrow = Cells(Rows.Count, "A").End(xlUp).Row

    '
    ' ' Column Headers
    Cells(i, 9).Value = "Stock Ticker Name"
    Cells(i, 11).Value = "Annual Percent Change"
    Cells(i, 12).Value = "Total Stock Volume"

     Summary_Table_Row = 2

      
    ' Start Loop through stock data 
        For i = 2 To Lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        ' For when the Stock Ticker Name changes
        If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value Then

 '    Finding the Values
            Stock_Ticker_Name = Cells(i, 1).Value
            Total_Stock_Volume = Cells(i, 7).Value
            Annual_Close_Price = Cells(i, 6).Value
            Annual_Open_Price = ells(i, 3).Value
            Annual_Change = Cells(i, 10).Value 

            Annual_Change = Annual_Close_Price - Annual_Open_Price

            Annual_Percent_Change = (Annual_Close_Price - Annual_Open_Price)/Annual_Close_Price

        'Print Values into Summary Table
            Range("I" & Summary_Table_Row).Value = Stock_Ticker_Name
            Range("J" & Summary_Table_Row).Value = Annual_Change
            Range("K" & Summary_Table_Row).Value = Annual_Percent_Change
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
         Summary_Table_Row = Summary_Table_Row + 1
            Total_Stock_Volume = 0

        ElseIf Cells(i, 10).Value > 0 Then
         Cells(i, 10).Interior.ColorIndex = 3  
            
        Else 
          ' Add to Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

       If Cells(1, 10).Value = < 0 Then
        Cells(1, 10).Interior.ColorIndex = 3

      End If  
    End If
  
  Next i


  Cells(1, 9).Font.Bold = True
  Cells(1, 10).Font.Bold = True
  Cells(1, 9).ColumnWidth = 20
  Cells(1, 10).ColumnWidth = 20

End Sub

