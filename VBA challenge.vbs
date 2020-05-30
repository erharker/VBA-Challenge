Sub VBAchallenge():
'variables
    Dim ticker As String
    Dim Yearly_change As Double
    Dim Percent_change As Double
    Dim total_stock_vol As Double
    total_stock_vol = 0
    Dim first_open As Double
    Dim final_close As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
              
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
               
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'The ticker symbol
            
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = ticker
            
        'total stock volume
            
        total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
        ws.Range("L" & Summary_Table_Row).Value = total_stock_vol
                            
        final_close = ws.Cells(i, 6).Value
        Yearly_change = final_close - first_open
        ws.Range("J" & Summary_Table_Row).Value = Yearly_change
        
        If first_open = 0 Then
            Percent_change = 0
        Else
            Percent_change = (Yearly_change / first_open)
        End If
        ws.Range("K" & Summary_Table_Row).Value = Percent_change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
        If Yearly_change > 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
            
        Summary_Table_Row = Summary_Table_Row + 1
        
        'column headers
            
        ws.Cells(1, 9).Value = "ticker"

        ws.Cells(1, 10).Value = "Yearly Change"

        ws.Cells(1, 11).Value = "Percent Change"

        ws.Cells(1, 12).Value = "Total Stock volume"
        
       
    
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        first_open = ws.Cells(i, 3).Value
    
    Else
    
        total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
                    
        End If
        
        
    
    Next i
    
Next ws
         

End Sub

