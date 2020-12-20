Sub stocksandtickers():

'Add Headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Loop through all the ticker sumbols

Dim ticker As String
Dim symbol As String
Dim Summary_tablerow As Long
Dim volume As LongLong
Dim i As Long
Dim open_price As Long
Dim close_price As Double
Dim change As Double
Dim percent_change As Double
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Activate
    
            
        'Add Headers
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        Summary_tablerow = 2
        volume = 0
        open_price = 0
        close_price = 0
        change = 0
        percent_change = 0

        'find last row
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        'compare one ticker to another ticker going to column A
        For i = 2 To LastRow
        'when they are not the same ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
              
                 'set ticker name
                ticker = Cells(i, 1).Value
                
                'open price grab
                open_price = Cells(i + 1, 3).Value
             
                'closing price grab
                closing_price = Cells(i, 6).Value
                
                'calculate yearly change
                change = closing_price - open_price
                
                'add yearly change to summary table
                Range("J" & Summary_tablerow) = change
    
                'get around 0 error
                If open_price = 0 Then
                    percent_change = 0
                    
                    Else
                    percent_change = Round(change - 1, 2)
                
                End If
                
                
                'color formatting for column K
                
                If change < 0 Then
                    Cells(Summary_tablerow, "K").Interior.ColorIndex = 3
                    
                 ElseIf change > 0 Then
                    Cells(Summary_tablerow, "K").Interior.ColorIndex = 4
                    
                Else
                    Cells(Summary_tablerow, "K").Interior.ColorIndex = 2
                    
                End If
                
                
                'add % change to summary table
                 Range("K" & Summary_tablerow) = percent_change
                
                'add to volume total
                volume = volume + Cells(i, 7).Value
                
                'print ticker in I column
                Range("I" & Summary_tablerow) = ticker
                
                'print volume to L column
                Range("L" & Summary_tablerow) = volume
                 
                 'add one to summary table for next ticker after adding ticker,ticker total
                Summary_tablerow = Summary_tablerow + 1
            
                
                'reset back to zero
                volume = 0
                
            'when they are the same ticker
            Else
               'add to volume for ticker
                volume = volume + Cells(i, 7).Value
     
            
            End If

    Next i
    
Next ws


End Sub

