'Dim starting_ws As Worksheet'
'Set starting_ws = ActiveSheet'
Dim totalVol As Double
Dim ticker As String
Dim yearStart As Double
Dim yearEnd As Double
Dim yearChange As Double
Dim yearPercent As Double
Dim c As Double
Dim v As Double

c = -1

totalVol = 0
yearPercent = 0
Dim tHeader, cHeader, pHeader, vHeader As String
tHeader = "Ticker"
cHeader = "Year Change"
pHeader = "Percent Change"
vHeader = "Total volume"

'Sets table headers'


'Iterates through each sheet in current workbook'
For Each ws In Worksheets

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'Sets table headers'
    Range("J1").Value = tHeader
    Range("K1").Value = cHeader
    Range("L1").Value = pHeader
    Range("M1").Value = vHeader

   
  'Gets last row in sheet'
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  'Iterates through all filled rows in the sheet'
  For i = 2 To lastRow
    
    'If the next ticker in the cell is not equal to the ticker previous cell'
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        'Sets ticker to the last cell not equal to previous cell'
        ticker = ws.Cells(i, 1).Value
        
        'Last addition to the total volume for each ticker'
        totalVol = totalVol + ws.Cells(i, 7).Value
        
        'Stores final closing value for each ticker'
        yearEnd = ws.Cells(i, 6).Value
        
        'Gets year changes by subtracting year open from year close'
        yearChange = yearEnd - yearStart
        
            'this block makes sure that no division by 0 occurs'
            If yearStart <> 0 Then
        
            yearPercent = (yearChange / yearStart) * 100
            
            Else
                yearPercent = 0
            End If
        
        'Prints ticker to table'
        ws.Range("J" & Summary_Table_Row) = ticker
        
        'Prints yearly price change to table'
        ws.Range("K" & Summary_Table_Row) = yearChange
        
            'Sets conditional formatting for negative and positive change values'
            If yearChange < 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            
            End If
        'Changes yearPercent to percent format and prints to table'
        ws.Range("L" & Summary_Table_Row) = Round(yearPercent, 2) & "%"
        
        'Prints total volume to table'
        ws.Range("M" & Summary_Table_Row) = totalVol
        
        'Adds to summary table row'
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Resets counter to zero'
        totalVol = 0
        'v = Cells(i + 1, 3).Value'
        
        'Range("R" & Summary_Table_Row) = v'
        
        'Resets iteration tracker to -1'
        c = -1
        
    Else
        'Adds to total volume'
        totalVol = totalVol + ws.Cells(i, 7).Value
        
        'Adds to loop tracker. Keeps count of how many iterations before ticker changes'
        c = c + 1
        'Gets first value (open) in ticker'
        yearStart = ws.Cells(i - c, 3).Value
        
    End If
    
    
    Next i
   
Next ws
   
End Sub

