# VBA-challenge

This vba code allows us to make a summary table of the data given calculating the
quarterly change, percent change, and total volume of each stock in the data for 
each quarter. Futhermore it allows to make another summary table in which the code 
finds the greatest percent increase in volume out of the stocks given as well as the 
greatest percent decrease and greatest volume.  

Some code were found using ChatGPT:
-finding the last row of the column
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
-formatting a column to be percent
ws.Range("K" & tickerRow).Value = Format(ws.Range("K" & tickerRow).Value, "Percent")
