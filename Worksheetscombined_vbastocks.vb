Sub vbastockss()


Dim Ws As Integer
For Ws = 1 To Sheets.Count
    Sheets(Ws).Activate


    'define variables, formatting and titles
    
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percentage Change per year"
    Range("M1").Value = "Total Volume"
    
    Range("P2").Value = "Greatest % increase"
    Range("P3").Value = "Greatest % decrease"
    Range("P4").Value = "Greatest total volume"
    
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    
    Range("R2:R3").NumberFormat = "0.00%"
    
    Dim yearclose As Integer
    Dim yearopen As Integer
    Dim volume As Integer
    volume = 0
    
    
    'zero the beginning for stats table
    Range("R2:R4") = 0
    
    
        'create row-counter i loop for column A and j loop for summary in column J
        i = 2
        j = 2
        
        'find opening value
        Cells(j, 12).Value = Cells(i, 3).Value
        'insert unique tickers
        Cells(j, 10) = Cells(i, 1).Value
        
        'loop till blank cell
        While Cells(i, 1).Value <> ""
            
                
            'if row does down equal row above
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'find closing value
            Cells(j, 11).Value = Cells(i, 6)
            'change in price
            Cells(j, 11).Value = Cells(j, 11).Value - Cells(j, 12).Value
            
            j = j + 1
            
            'insert unique tickers
            Cells(j, 10) = Cells(i + 1, 1).Value
            'find opening value
            Cells(j, 12).Value = Cells(i + 1, 3).Value
            
            End If
            'sum the total volume
            Cells(j, 13).Value = Cells(j, 13).Value + Cells(i, 7).Value
           
        i = i + 1
        Wend
    
            K = 2
            'loop through column 11 (unique tickers), and divide by column 12 - replace into column 12
             While Cells(K, 10).Value <> ""
             
             'calc the % change
                   
             Cells(K, 12).Value = Cells(K, 11).Value / Cells(K, 12).Value
             
             'conditional formatting column j
             
             If Cells(K, 11).Value > 0 Then
             Cells(K, 11).Interior.ColorIndex = 4
             ElseIf Cells(K, 11).Value < 0 Then
             Cells(K, 11).Interior.ColorIndex = 3
             End If
             
             'change to percentage
             Cells(K, 12).NumberFormat = "0.00%"
             
             'Challenge
                    
            'Greatest % increase
             If Cells(2, 18).Value < Cells(K, 12).Value Then
             Cells(2, 18).Value = Cells(K, 12).Value
             Cells(2, 17).Value = Cells(K, 10)
             End If
             
             'Greatest % decrease
             If Cells(3, 18).Value > Cells(K, 12).Value Then
             Cells(3, 18).Value = Cells(K, 12).Value
             Cells(3, 17).Value = Cells(K, 10)
             End If
             
              'Greatest volume
             If Cells(4, 18).Value < Cells(K, 13).Value Then
             Cells(4, 18).Value = Cells(K, 13).Value
             Cells(4, 17).Value = Cells(K, 10)
             End If
             
             K = K + 1
             Wend
Next Ws

End Sub
    



