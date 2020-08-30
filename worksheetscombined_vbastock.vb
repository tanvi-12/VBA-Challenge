Sub vbastockss()

'Note this is the version with worksheets

For Each ws In Worksheets

    'define variables, formatting and titles

    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percentage Change per year"
    ws.Range("M1").Value = "Total Volume"

    ws.Range("P2").Value = "Greatest % increase"
    ws.Range("P3").Value = "Greatest % decrease"
    ws.Range("P4").Value = "Greatest total volume"

    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"

    ws.Range("R2:R3").NumberFormat = "0.00%"

    Dim yearclose As Integer
    Dim yearopen As Integer
    Dim volume As Integer
    volume = 0


    'zero the beginning for stats table
    ws.Range("R2:R4") = 0


        'create row-counter i loop for column A and j loop for summary in column J
        i = 2
        j = 2
        
        'find opening value
        ws.Cells(j, 12).Value = Cells(i, 3).Value
        'insert unique tickers
        ws.Cells(j, 10) = Cells(i, 1).Value
        
        'loop till blank cell
        While ws.Cells(i, 1).Value <> ""
            
            'sum the total volume
            ws.Cells(j, 13).Value = ws.Cells(j, 13).Value + ws.Cells(i, 7).Value
            'if row does down equal row above
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'find closing value
            ws.Cells(j, 11).Value = ws.Cells(i, 6)
            'change in price
            ws.Cells(j, 11).Value = ws.Cells(j, 11).Value - ws.Cells(j, 12).Value
            
            j = j + 1
            
            'insert unique tickers
            ws.Cells(j, 10) = ws.Cells(i + 1, 1).Value
            'find opening value
            ws.Cells(j, 12).Value = ws.Cells(i + 1, 3).Value
            
            End If
        
        
        i = i + 1
        Wend

            K = 2
            'loop through column 11 (unique tickers), and divide by column 12 - replace into column 12
            While ws.Cells(K, 10).Value <> ""
            
            'calc the % change
                
            ws.Cells(K, 12).Value = ws.Cells(K, 11).Value / ws.Cells(K, 12).Value
            
            'conditional formatting column j
            
            If ws.Cells(K, 11).Value > 0 Then
            ws.Cells(K, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(K, 11).Value < 0 Then
            ws.Cells(K, 11).Interior.ColorIndex = 3
            End If
            
            'change to percentage
            ws.Cells(K, 12).NumberFormat = "0.00%"
            
            'Challenge
                    
            'Greatest % increase
            If ws.Cells(2, 18).Value < ws.Cells(K, 12).Value Then
            ws.Cells(2, 18).Value = ws.Cells(K, 12).Value
            ws.Cells(2, 17).Value = ws.Cells(K, 10)
            End If
            
            'Greatest % decrease
            If ws.Cells(3, 18).Value > ws.Cells(K, 12).Value Then
            ws.Cells(3, 18).Value = ws.Cells(K, 12).Value
            ws.Cells(3, 17).Value = ws.Cells(K, 10)
            End If
            
            'Greatest volume
            If ws.Cells(4, 18).Value < ws.Cells(K, 13).Value Then
            ws.Cells(4, 18).Value = ws.Cells(K, 13).Value
            ws.Cells(4, 17).Value = ws.Cells(K, 10)
            End If
            
            K = K + 1
            Wend

Next ws
 
End Sub