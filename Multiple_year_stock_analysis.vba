' The purpose of this code is to quickly go through multiple Excel sheets of yearly
' stock ticker data and provide easily digestible yearly stock performance for each
' stock.  The code calculates the yearly stock change along with conditional 
' formating (green = + , red = -) based on increase or decrease in stock value. the
' code also calculates the percent change over the year and gives total yearly
' volume of the stock.  It then calculates the stocks with greatest increase, 
' decrease, and volume over the year.

Sub StockAnalysis()

dim ws as Worksheet
dim tick as string
dim int_value as double
dim end_value as double
dim lastrow as long
dim lr_ticks as long
dim t as string
dim start_value as long
dim StartRow as long
dim EndRow as long
dim yearly_change as double
dim percent_change as double
dim yearly_volume as double
dim greatest as double
dim least as double
dim GVol as double
dim G_row as long
dim L_row as long
dim GV_row as long



' loops through all the worksheets in the workbook
for each ws in Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
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
    start_value = 2
    ' loops through tick for each ticker value
    for i = 3 to lr_ticks

        ' places the value from each cell into the tick variable.  
        ' there is a duplicate of the first tick so the loop
        ' starts at row 3 skipping the first instance of the first
        ' tick value. to correct the duplicate value the rows 3 to
        ' last row are shifted up and the contents of the last row are
        ' cleared.
        if (i <> lr_ticks) then
            tick = ws.Cells(i, 9).value
            ws.Cells(i-1, 9).Value = tick
        else
            tick = ws.Cells(i, 9).value
            ws.Cells(i-1, 9).Value = tick
            ws.Cells(i, 9).ClearContents
        end if

        StartRow = start_value

        for j = start_value to lastrow

            if (ws.Cells(j, 1) <> tick) then
                
                EndRow = j-1
                start_value = j
                exit for

            elseif (j = lastrow) then

                EndRow = j

            end if
        next j
        
        for k = StartRow to EndRow
            if (ws.Cells(k, 3).Value <> 0) then
                int_value = ws.Cells(k, 3).Value
                exit for
            End if
        next k
        
        end_value = ws.Cells(EndRow, 6).Value
        
        yearly_change = end_value - int_value
        
        if (yearly_change = 0) then
            percent_change = 0
        else
            percent_change = (yearly_change / int_value) * 100
        End if
        yearly_volume = 0
        for n = StartRow to EndRow

            yearly_volume = yearly_volume + ws.Cells(n,7)

        next n

        ws.Cells(i-1, 10).Value = yearly_change

        if yearly_change > 0 then
            ws.Cells(i-1, 10).Interior.ColorIndex = 4
        else
            ws.Cells(i-1, 10).Interior.ColorIndex = 3
        End if
        ws.Cells(i-1, 11).Value = percent_change
        ws.Cells(i-1, 12).Value = yearly_volume
    next i
    
    greatest = Cells(2, 11).Value
    least = Cells(2, 11).Value
    GVol = Cells(2, 12).Value
    G_row = 2
    L_row = 2
    GV_row = 2
    for m = 3 to (lr_ticks - 1)

        if (Cells(m, 11).Value > greatest) then
            greatest = Cells(m, 11).Value
            G_row = m
        End if

        if (Cells(m, 11).Value < least) then
            least = Cells(m, 11).Value
            L_row = m
        End if

        if (Cells(m, 12).Value > GVol) then
            GVol = Cells(m, 12).Value
            GV_row = m
        End if

    next m

    ' insert ticker and value for greatest % increase
    Cells(2, 15).Value = Cells(G_row, 9).Value
    Cells(2, 16).Value = greatest

    ' insert ticker and value for greatest % decrease
    Cells(3, 15).Value = Cells(L_row, 9).Value
    Cells(3, 16).Value = least

    ' insert ticker and value for greatest total volume
    Cells(4, 15).Value = Cells(GV_row, 9).Value
    Cells(4, 16).Value = GVol
    

next ws


End Sub