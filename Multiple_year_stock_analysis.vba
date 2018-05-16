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
'dim int_value as float
'dim end_value as float
dim lastrow as long
dim lr_ticks as long
dim t as string

' loops through all the worksheets in the workbook
for each ws in Worksheets

    '-------------------------------------------------------------
    ' extracts the last row from collumn 1 (A)
    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    '------------------------------------------------------------
    ' Activates the worksheet in ws so that in the next step
    ' Active sheet can be used in the advancedfilter method.
    ws.Activate
    '------------------------------------------------------------
    ' extracts the unique values from the ticker collumn
    ' and inserts them to column I. Causes an extra value of
    ' the first ticker value to be inserted.
    ActiveSheet.Range("A2:A" & lastrow).AdvancedFilter _
    Action:=xlFilterCopy, _
    CopyToRange:=ActiveSheet.Range("I2"), _
    Unique:=True
    '-------------------------------------------------------------
    ' finds the last row of column I containing the new unique tick
    ' values.
    lr_ticks = ws.Cells(Rows.count, 9).End(xlUp).Row
    
    '-------------------------------------------------------------
    ' loops through tick for each ticker value
    for i = 3 to lr_ticks

        ' places the value from each cell into the array "tick".  
        ' there is a duplicate of the first tick so the loop
        ' starts at row 3 skipping the first instance of the first
        ' tick value.
        tick(i - 3) = ws.Cells(i, 9).value

    next i
    ws.columns(9).ClearContents
    Range(ws.Cells(2,9), ws.Cells(lr_ticks-1,9)).Value = tick

next ws


End Sub