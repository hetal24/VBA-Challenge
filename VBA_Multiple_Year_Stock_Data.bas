Attribute VB_Name = "Module1"
Sub stockData()

    ' Create the variable
    Dim out_row, row1, voume  As Integer
    Dim o_price, c_price, g_increase, g_decrease, g_volume As Double
    Dim ticker As String
    


    ' Loop thtough all sheets.
    For Each ws In ThisWorkbook.Worksheets
    
        ' Determin the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    
        ' Lable the column
        ws.Cells(1, "j").Value = "Ticker"
        ws.Cells(1, "k").Value = "Quarterly change"
        ws.Cells(1, "L").Value = "Percentage Chnage"
        ws.Cells(1, "M").Value = "Total Stock Volume"
        ws.Cells(1, "Q").Value = "Ticker"
        ws.Cells(1, "R").Value = "Value"
        ws.Cells(2, "p").Value = "Greatest % Increase"
        ws.Cells(3, "p").Value = "Greatest % Decrease"
        ws.Cells(4, "p").Value = "Greatest Total Volume"
        
        ' set the variable default value
        out_row = 2
        row1 = 1
        volume = 0
        g_increase = 0
        g_decrease = 0
        g_volume = 0
        
        ' Loop through all row of data
        For r = 2 To lastrow
            If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
            
                ticker = ws.Cells(r, 1).Value
                ws.Cells(out_row, "j").Value = ticker
                
                c_price = ws.Cells(r, "f").Value
                o_price = ws.Cells(row1 + 1, "c").Value
                
                'For Quarterly change
                ws.Cells(out_row, "k").Value = c_price - o_price
                
                'For Percentage change
                
                ws.Cells(out_row, "L").Value = ((c_price - o_price) / o_price) * 100 & "%"
                
                ' Set cell color according to their value
                If ws.Cells(out_row, "k").Value < 0 Then
                    ws.Cells(out_row, "k").Interior.ColorIndex = 3
                End If
                If ws.Cells(out_row, "k").Value > 0 Then
                    ws.Cells(out_row, "k").Interior.ColorIndex = 4
                End If
              
               
                'For Total Stock Volume
                ws.Cells(out_row, "M").Value = volume + ws.Cells(r, "G").Value
                
               
                'Check Greatest % increase, Greatest % decrease, and Greatest total volume of the ticker
                'And their value
                
                If ws.Cells(out_row, "L").Value > g_increase Then
                    g_increase = ws.Cells(out_row, "L").Value
                   
                    ws.Cells(2, "Q").Value = ticker
                
                End If
                If ws.Cells(out_row, "L").Value < g_decrease Then
                    g_decrease = ws.Cells(out_row, "L").Value
                    
                    ws.Cells(3, "Q").Value = ticker
                End If
                
                If ws.Cells(out_row, "m").Value > g_volume Then
                    g_volume = ws.Cells(out_row, "m").Value
                    ws.Cells(4, "Q").Value = ticker
                    
                End If
                
               out_row = out_row + 1
               volume = 0
               row1 = r
            Else
                volume = volume + Cells(r, "G").Value
            End If
    
    Next r
    
     'Assign value into the cell
     ws.Cells(2, "R").Value = g_increase * 100 & "%"
     ws.Cells(3, "R").Value = g_decrease * 100 & "%"
     ws.Cells(4, "R").Value = g_volume
     
 Next ws

End Sub


