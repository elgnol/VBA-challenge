Attribute VB_Name = "Module1"
Sub year_stock()
    
    'Loop through all worksheets
    For Each ws In Worksheets
    
        Dim ticker As String
        Dim volume As Double
        Dim lowPercent As Double
        Dim greatPercent As Double
        Dim greatV As Double
        Dim lastRow As Long
        Dim tickerRow As Integer
        Dim openStart As Long
        
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        volume = 0
        tickerRow = 2
        openStart = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To lastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker = Cells(i, 1).Value
                volume = volume + ws.Cells(i, 7).Value
                
                'set ticket symbol in ticker column
                ws.Range("I" & tickerRow).Value = ticker
                
                'set total volume in total volume column
                ws.Range("L" & tickerRow).Value = volume
                
                'calculate quarterly change value
                ws.Range("J" & tickerRow).Value = ws.Cells(i, 6).Value - ws.Cells(openStart, 3).Value
                
                'Format the color of the quarterly change column
                'if it's positive it will be green if negative
                'it will be red and if 0 it will be left alone
                If ws.Range("J" & tickerRow).Value > 0 Then
                
                    ws.Range("J" & tickerRow).Interior.ColorIndex = 4
                
                ElseIf ws.Range("J" & tickerRow).Value < 0 Then
                
                    ws.Range("J" & tickerRow).Interior.ColorIndex = 3
                
                End If
                
                
                'calculate value of percent change
                If ws.Range("J" & tickerRow).Value <> 0 Then
                
                    ws.Range("K" & tickerRow).Value = ws.Range("J" & tickerRow).Value / ws.Cells(openStart, 3).Value
                
                Else
                
                    ws.Range("K" & tickerRow).Value = 0
                
                End If
                
                'Format percent change column to percent
                ws.Range("K" & tickerRow).Value = Format(ws.Range("K" & tickerRow).Value, "Percent")
                
                'move the ticker row to the next ticker
                tickerRow = tickerRow + 1
                
                'move to next ticker open price row
                openStart = i + 1
                
                'set volume value holder back to 0
                volume = 0
            
            Else
                
                'add up volume of the ticker
                volume = volume + ws.Cells(i, 7).Value
            
            End If
            
        
        Next i
        
        'functionality to return Greatest% increase, Greatest% decrease
        'and Greatest total volume
        Dim nLastRow As Long
        
        'track last row for the previously created ticker column
        nLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        lowPercent = 0
        greatPercent = 0
        greatV = 0
        
        'loop to find greatest % increase, greatest % decrease
        'and greatest volume
        For i = 2 To nLastRow
            
            If ws.Range("K" & i).Value < 0 Then
                            
                'test if percent change value is less than
                'current lowest percent holder
                'if so change new lowest percent holder to current
                'index percent change value
                If ws.Range("K" & i).Value < lowPercent Then
                    
                    lowPercent = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                    ws.Range("Q3").Value = lowPercent
                    
                End If
                
            
            Else
                
                'test if percent change value is greater than
                'current highest percent holder
                'if so change new highest percent holder to current
                'index percent change value
                If ws.Range("K" & i).Value > greatPercent Then
                
                    greatPercent = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                    ws.Range("Q2").Value = greatPercent
                    
                End If
                
            
            End If
            
            'format appropriate columns to percent
            ws.Range("Q2").Value = Format(ws.Range("Q2").Value, "Percent")
            ws.Range("Q3").Value = Format(ws.Range("Q3").Value, "Percent")
            
            'test if current index's stock volume value is greater
            'than the greatest volume place holder
            'if so change the greatest volume place holder
            'equals to current index's stock volume value
            If ws.Range("L" & i).Value > greatV Then
            
                greatV = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = greatV
            
            End If
            
            
        Next i
        
    
    Next ws
    

End Sub
