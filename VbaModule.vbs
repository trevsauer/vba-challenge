Attribute VB_Name = "Module1"
Sub Stocks()
    
    ' loop through every work sheet and apply the for loop
    For Each ws In Worksheets
    
        ' loops through all worksheets from beginning to end
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' set column header names to each of the columns in the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ' make values being calculated doubles
        Dim Yearly_Change As Double
        Dim Yearly_Start As Double
        Dim Yearly_End As Double
        Dim Start_Row As Double
        Dim Percent_Change As Double
        
        ' set the individual ticker name as a string
        Dim Ticker_Name As String
    
        ' set the Strock volume, collective stocks per ticker name, as a double
        Dim Stock_Volume As Double
        ' beginning value of the stock total is 0, sum of ticker name values added to 0
        Stock_Volume = 0
        
        'make counter to loop through tickers and start over once hits new ticker
        Dim Row_Count As Double
        Row_Count = 0
    
        ' create variable for summary table row, second row aka column B
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
        
        Dim i As Long
        
        For i = 2 To LastRow
            
            ' only loops if values arent equal to eachother
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            ' places tickers in summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ' calculate yearly change
            ' If yearly start and end are = 0, then set summary table values to 0
                
            Start_Row = i - Row_Count
            
            ' if loop accounts for any zeros found, since dividing by a 0 gives an error
            If ws.Cells(i, 6).Value Or ws.Range("C" & Start_Row).Value <> 0 Then
                
                ' calculate yearly change
                Yearly_End = ws.Cells(i, 6).Value
                Yearly_Start = ws.Range("C" & Start_Row).Value
                Yearly_Change = Yearly_Start - Yearly_End
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                ' calculate percent change
                
                Percent_Change = Yearly_Change / Yearly_End * 100
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                            
                End If
                
                    ' calculate stock volume
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                    ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                
                    ' add to counter and stock volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    Row_Count = 0
                    Stock_Volume = 0
                        
                Else
                    
                    Row_Count = Row_Count + 1
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                     
                End If
                
                    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                        ' makes positive cells green
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                        ' makes negative cells red
                End If
            
            Next i

    Next ws

    
End Sub

