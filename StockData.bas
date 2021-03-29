Attribute VB_Name = "Module1"
Sub vba_challenge_2():

    Dim Ticker As String
        Ticker = i
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
    Dim Summary_Table_Row As Double
    Dim LastRow As Double
    
        
    For Each ws In Worksheets
        
        Summary_Table_Row = 2
        
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
    ws.Activate
    
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Opening_Price"
        ws.Cells(1, "K").Value = "Closing_Price"
        ws.Cells(1, "L").Value = "Yearly_Change"
        ws.Cells(1, "M").Value = "Percent_Change"
        ws.Cells(1, "N").Value = "Total_Stock_Volume"
        
        
    For i = 2 To LastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, "A").Value
            Debug.Print (Ticker)
            
            Opening_Price = ws.Cells(i, "C").Value
            Debug.Print (Opening_Price)
            
            Closing_Price = ws.Cells(i, "F").Value
            Debug.Print (Closing_Price)
            
            Yearly_Change = Closing_Price - Opening_Price
            Debug.Print (Yearly_Change)
            
            If Opening_Price = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / Opening_Price
            Debug.Print (Percent_Change)
            End If
                                                     
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
            Debug.Print (Total_Stock_Volume)
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ws.Range("J" & Summary_Table_Row).Value = Opening_Price
            
            ws.Range("K" & Summary_Table_Row).Value = Closing_Price
            
            ws.Range("L" & Summary_Table_Row).Value = Yearly_Change
                            
               If ws.Range("L" & Summary_Table_Row).Value < 0 Then
        
                  ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                    
               Else
                  ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
               Debug.Print
               End If
               
            ws.Range("M" & Summary_Table_Row).Value = Percent_Change
                ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"
                
            ws.Range("N" & Summary_Table_Row).Value = Total_Stock_Volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Total_Stock_Volume = 0
            
        Else
        
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
            
        End If
        
                
    Next i
    
    Next ws
       
        
    
End Sub


