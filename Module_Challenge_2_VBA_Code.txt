Option Explicit
Sub Alphabet_Ticker_Loop()
    Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate
    
    Const TICKER_COLUMN As Integer = 1
    Const OPENING_PRICE_COLUMN As Integer = 3
    Const CLOSING_PRICE_COLUMN As Integer = 6
    Const STOCK_VOLUME_COLUMN As Integer = 7
    Const UNIQUE_TICKER_COLUMN As Integer = 9
    Const YEARLY_CHANGE_COLUMN As Integer = 10
    Const PERCENTAGE_CHANGE_COLUMN As Integer = 11
    Const TOTAL_STOCK_VOL_COLUMN As Integer = 12
    Const GREATEST_TICKER_COLUMN As Integer = 16
    Const GREATEST_VALUE_COLUMN As Integer = 17
    Const FIRST_DATA_ROW As Integer = 2
    
    Dim Current_Ticker As String
    Dim Next_Ticker As String
    Dim Total_Stock_Vol As LongLong
    Dim Opening_Value As Double
    Dim Closing_Value As Double
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Output_Row As Integer
    Dim lastrow As Long
    Dim Input_Row As Long
    Dim Max_Value_Ticker As String
    Dim Max_Value As Double
    Dim Min_Value_Ticker As String
    Dim Min_Value As Double
    Dim Max_Total_Stock_Vol_Ticker As String
    Dim Max_Total_Stock_Vol As Double
    
    
    Total_Stock_Vol = 0
    Opening_Value = Cells(FIRST_DATA_ROW, OPENING_PRICE_COLUMN).Value
    Output_Row = FIRST_DATA_ROW
    lastrow = Cells(Rows.Count, TICKER_COLUMN).End(xlUp).Row

    
    For Input_Row = FIRST_DATA_ROW To lastrow
        Current_Ticker = Cells(Input_Row, TICKER_COLUMN).Value
        Next_Ticker = Cells(Input_Row + 1, TICKER_COLUMN).Value
        Total_Stock_Vol = Total_Stock_Vol + Cells(Input_Row, STOCK_VOLUME_COLUMN).Value
       
        If Next_Ticker <> Current_Ticker Then
            'Inputs
            Closing_Value = Cells(Input_Row, CLOSING_PRICE_COLUMN).Value
            
            'Calculations
            Yearly_Change = Closing_Value - Opening_Value
            Percentage_Change = (Yearly_Change / Opening_Value)
            
            'Outputs
            'MsgBox (Current_Ticker & " " & Yearly_Change & " " & Percentage_Change)
            Cells(1, UNIQUE_TICKER_COLUMN).Value = "Ticker"
            Cells(Output_Row, UNIQUE_TICKER_COLUMN).Value = Current_Ticker
            
            Cells(1, YEARLY_CHANGE_COLUMN).Value = "Yearly Change"
            Cells(Output_Row, YEARLY_CHANGE_COLUMN).Value = Yearly_Change
            
            Cells(1, PERCENTAGE_CHANGE_COLUMN).Value = "Percentage Change"
            Cells(Output_Row, PERCENTAGE_CHANGE_COLUMN).Value = Percentage_Change
            Cells(Output_Row, PERCENTAGE_CHANGE_COLUMN).NumberFormat = "0.00%"
            
            Cells(1, TOTAL_STOCK_VOL_COLUMN).Value = "Total Stock Volume"
            Cells(Output_Row, TOTAL_STOCK_VOL_COLUMN).Value = Total_Stock_Vol
            
            If Yearly_Change >= 0 Then
                Cells(Output_Row, YEARLY_CHANGE_COLUMN).Interior.ColorIndex = 4
            
            ElseIf Yearly_Change <= 0 Then
                Cells(Output_Row, YEARLY_CHANGE_COLUMN).Interior.ColorIndex = 3
            
            End If
                
            'Set Up for Next Row
            Output_Row = Output_Row + 1
            Total_Stock_Vol = 0
            Opening_Value = Cells(Input_Row + 1, OPENING_PRICE_COLUMN).Value
        End If
    Next Input_Row
    
    'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
    Max_Value_Ticker = " "
    Max_Value = 0
    Min_Value_Ticker = " "
    Min_Value = 0
    Max_Total_Stock_Vol_Ticker = " "
    Max_Total_Stock_Vol = 0
   
   lastrow = Cells(Rows.Count, UNIQUE_TICKER_COLUMN).End(xlUp).Row
    
    For Input_Row = FIRST_DATA_ROW To lastrow
    
        If Cells(Input_Row, PERCENTAGE_CHANGE_COLUMN).Value > Max_Value Then
        'Inputs
            Max_Value = Cells(Input_Row, PERCENTAGE_CHANGE_COLUMN).Value
            Max_Value_Ticker = Cells(Input_Row, UNIQUE_TICKER_COLUMN).Value
        End If
          
          If Cells(Input_Row, PERCENTAGE_CHANGE_COLUMN).Value < Min_Value Then
        'Inputs
            Min_Value = Cells(Input_Row, PERCENTAGE_CHANGE_COLUMN).Value
            Min_Value_Ticker = Cells(Input_Row, UNIQUE_TICKER_COLUMN).Value
        End If
         
         If Cells(Input_Row, TOTAL_STOCK_VOL_COLUMN).Value > Max_Total_Stock_Vol Then
        'Inputs
            Max_Total_Stock_Vol = Cells(Input_Row, TOTAL_STOCK_VOL_COLUMN).Value
            Max_Total_Stock_Vol_Ticker = Cells(Input_Row, UNIQUE_TICKER_COLUMN).Value
        End If
    
    Next Input_Row
        'Outputs
            Cells(1, GREATEST_TICKER_COLUMN).Value = "Ticker"
            Cells(1, GREATEST_TICKER_COLUMN).Value = "Value"
            
            Cells(2, "O").Value = "Greatest % Increase"
            Cells(2, GREATEST_TICKER_COLUMN).Value = Max_Value_Ticker
            Cells(2, GREATEST_VALUE_COLUMN).Value = Max_Value
            Cells(2, GREATEST_VALUE_COLUMN).NumberFormat = "0.00%"
            
            Cells(3, "O").Value = "Greatest % Decrease"
            Cells(3, GREATEST_TICKER_COLUMN).Value = Min_Value_Ticker
            Cells(3, GREATEST_VALUE_COLUMN).Value = Min_Value
            Cells(3, GREATEST_VALUE_COLUMN).NumberFormat = "0.00%"
            
            Cells(4, "O").Value = "Greatest Total Volume"
            Cells(4, GREATEST_TICKER_COLUMN).Value = Max_Total_Stock_Vol_Ticker
            Cells(4, GREATEST_VALUE_COLUMN).Value = Max_Total_Stock_Vol
            
    Next ws
    
    MsgBox ("Done")
End Sub



