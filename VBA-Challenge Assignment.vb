Sub stockdata()
    ' run the same below code until it runs on every workbook
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    Dim ticker As String
    Dim i As Long
    Dim volume As Double
    Dim lastRow As Long
    Dim summary_table_row As Long
    Dim quarterlychange As Double
    Dim percentchange As Double
    Dim openvalue As Double
    Dim closevalue As Double
    Dim maxincrease As Double
    Dim minincrease As Double
    Dim maxvolincrease As Double
    Dim maxticker As String
    Dim minticker As String
    Dim maxvolticker As String
        
    ' Set headers for summary tables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Changed"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    
    ' Find the last row of data in column A
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Initialize summary table row
    summary_table_row = 2

    ' Initialize ticker volume

    volume = 0

    ' Initialize openvalue to select the first value in column "C"

    openvalue = ws.Cells(2, 3).Value

    ' initialize min & max increase
    
    maxincrease = 0
    minincrease = 1000 'randomly chose 1000 (the higher number) to avoid conveying wrong result
    maxvolincrease = 0
  
    For i = 2 To lastRow
        ' Check if the next row's ticker symbol is different
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' If different, update ticker, volume, and summary table
            ticker = ws.Cells(i, 1).Value
            volume = volume + ws.Cells(i, 7).Value
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("L" & Summary_Table_Row).Value = volume
            
            closevalue = ws.Cells(i, 6).Value
            
            quarterlychange = closevalue - openvalue
            percentchange = quarterlychange / openvalue

            ws.Range("J" & Summary_Table_Row).Value = quarterlychange
            ws.Range("K" & Summary_Table_Row).Value = percentchange

            'iterate the sequence in the column K to find the the least increase or greatest increase %
            'iterate the sequence in the column L to find the maximum volume increase

            If percentchange > maxincrease Then
            maxincrease = percentchange
            maxticker = ticker
            End If
                        
            If percentchange < minincrease Then
            minincrease = percentchange
            minticker = ticker
            End If
                        
            If volume > maxvolincrease Then
            maxvolincrease = volume
            maxvolticker = ticker
            End If

    
            'formting the cells

            ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'coloring the column "J", but no color for 0.00 (no change value)

            If quarterlychange < 0 Then
            
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            ElseIf quarterlychange > 0 Then
            
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                                   
            End If
            
                                
            Summary_Table_Row = Summary_Table_Row + 1
            ' Reset volume, quarterlychange, percentchange for the next ticker
            ' Redefine openvalue for the next ticker
            volume = 0
            quarterlychange = 0
            percentchange = 0
            openvalue = ws.Cells(i + 1, 3).Value

        Else
            ' If same ticker, continue adding volume
            volume = volume + ws.Cells(i, 7).Value
        End If

    Next i
    
    ws.Range("Q2").Value = maxincrease
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ws.Range("Q3").Value = minincrease
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Range("Q4").Value = maxvolincrease
    
    ws.Range("P2").Value = maxticker
    ws.Range("P3").Value = minticker
    ws.Range("P4").Value = maxvolticker
    
    
    
Next ws

End Sub

