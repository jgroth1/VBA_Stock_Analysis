' The purpose of this code is to quickly go through multiple Excel sheets of yearly
' stock ticker data and provide easily digestible yearly stock performance for each
' stock.  The code calculates the yearly stock change along with conditional 
' formating (green = + , red = -) based on increase or decrease in stock value. the
' code also calculates the percent change over the year and gives total yearly
' volume of the stock.  It then calculates the stocks with greatest increase, 
' decrease, and volume over the year.

Sub StockAnalysis()
dim ws as Worksheet
dim tick() as string
dim int_value as float
dim end_value as float
dim lastrow as long
dim lr_ticks as long
dim t as string

' loops through all the worksheets in the workbook
for each ws in Worksheets

    '-------------------------------------------------------------
    ' extracts the last row from collumn 1
    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    '------------------------------------------------------------
    ' Activates the worksheet in ws
    ws.Activate
    '------------------------------------------------------------
    ' extracts the unique values from the ticker collumn
    ' and inserts them to collumn J
    ActiveSheet.Range("A2:A" & lastrow).AdvancedFilter _
    Action:=xlFilterCopy, _
    CopyToRange:=ActiveSheet.Range("I2"), _
    Unique:=True
    '-------------------------------------------------------------
    ' places the unique ticker values into array tick
    lr_ticks = ws.Cells(Rows.count, 9).End(xlUp).Row
    tick = Range(ws.Cells(3,9), ws.Cells(lr_ticks,9)).Value
    '-------------------------------------------------------------
    ' loops through tick for each ticker value
    for each t in tick
        
        

    next
    

next ws


End Sub