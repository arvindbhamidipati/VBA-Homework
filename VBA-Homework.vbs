Sub sheetLoop():
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Select
        Call Assignment
        
    Next
End Sub

Sub Assignment():
    
    Dim Ticker_name As String
    Dim lRow As Long
    Dim lCol As Long
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim Open_n As Double
    Dim Close_n As Double
    Dim Percent_Changed As Double
    Dim Stock_Volume As Double
    
    Stock_Volume = 0
    
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    
    'Count for rows and columns stored in variables
    'Got the structure from https://www.excelcampus.com/vba/find-last-row-column-cell/
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Going through the full ticker
    For i = 2 To lRow
        'Need to calculate the first open number per group
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            Open_n = Cells(i, 3).Value
        End If
        
        'Need to grab each unique Ticker name, Closing number, and calculate the difference between the 2
        'Increment the for the summary table to the right
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Ticker_name = Cells(i, 1).Value
            Close_n = Cells(i, 6)
            
            Range("I" & Summary_Table_Row).Value = Ticker_name
            
            'Y0early Change
            Cells(Summary_Table_Row, 10).Value = Close_n - Open_n
            If Close_n - Open_n < 0 Then
                Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            Else
                Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            End If
            
            
            'Percent Changed
            If (Open_n < 0 Or Open_n = 0) Then
                Percent_Changed = 0
            Else
                Percent_Changed = (Close_n - Open_n) / Open_n
                Cells(Summary_Table_Row, 11).Value = Format(Percent_Changed, "#.##%")
            End If
            
            
            'Total Stock Volume
            Stock_Volume = Stock_Volume + Cells(i, 7)
            Cells(Summary_Table_Row, 12).Value = Stock_Volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            Stock_Volume = 0
        Else
            'Add to the Stock_Volume
            Stock_Volume = Stock_Volume + Cells(i, 7)
            Cells(Summary_Table_Row, 12).Value = Stock_Volume
            
        End If
    Next i
    
End Sub