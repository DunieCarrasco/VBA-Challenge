Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock()
    
'Define the variables

    Dim Total_Volumn As Double
    Dim Open_Beginning As Double
    Dim Close_End As Double
    
'Define variables for loops
    Dim i As Integer
    Dim j As Long
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    
    Dim ws_count As Integer
    
    k = 2
    
    ws_count = ActiveWorkbook.Worksheets.Count
    
' Loops through all sheets
    For i = 1 To ws_count
' Labels collected data columns
        ActiveWorkbook.Worksheets(i).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 12).Value = "Total Stock Volume"
'Lables increase/decrease and greatest volumn table
        ActiveWorkbook.Worksheets(i).Cells(1, 16).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 17).Value = "Value"
        ActiveWorkbook.Worksheets(i).Cells(2, 15).Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(i).Cells(3, 15).Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(i).Cells(4, 15).Value = "Greatest Total Volume"
        
        
        Total_Volumn = 0
        
        
'Looping entire sheet
        For j = 2 To ActiveWorkbook.Worksheets(i).Cells.SpecialCells(xlCellTypeLastCell).Row
            If ActiveWorkbook.Worksheets(i).Cells(j, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j + 1, 1).Value Then
            
            
'Loop for finding ticker
                ActiveWorkbook.Worksheets(i).Cells(k, 9).Value = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
'Loop for finding value and percent difference
                Close_End = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                ActiveWorkbook.Worksheets(i).Cells(k, 10).Value = Close_End - Open_Beginning
                ActiveWorkbook.Worksheets(i).Cells(k, 11).Value = (Close_End - Open_Beginning) / Open_Beginning
                    Open_Beginning = 0
                    Close_End = 0
'Loop for finding total volumn
                Total_Volumn = Total_Volumn + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                ActiveWorkbook.Worksheets(i).Cells(k, 12).Value = Total_Volumn
                    Total_Volumn = 0
                     k = k + 1
                
            ElseIf ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j, 1).Value Then
'Loop for finding value and percent difference
                Open_Beginning = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value
'Loop for finding total_volumn
                Total_Volumn = Total_Volumn + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            Else
                Total_Volumn = Total_Volumn + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            End If
        Next j
'Formating color change for negative and positive change
        For l = 2 To ActiveWorkbook.Worksheets(i).Range("J1").CurrentRegion.Rows.Count
            ActiveWorkbook.Worksheets(i).Cells(l, 11).Style = "Percent"
            If ActiveWorkbook.Worksheets(i).Cells(l, 10).Value > 0 Then
                With ActiveWorkbook.Worksheets(i).Cells(l, 10).Interior
                .ColorIndex = 4
                End With
            Else
                With ActiveWorkbook.Worksheets(i).Cells(l, 10).Interior
                .ColorIndex = 3
                   
                End With
            End If
        Next l
        
'Finding greatest/lowest increase in percent and greatest total volumn
    ActiveWorkbook.Worksheets(i).Cells(2, 17).Value = WorksheetFunction.Max(Worksheets(i).Range("L2:L" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
    ActiveWorkbook.Worksheets(i).Cells(3, 17).Value = WorksheetFunction.Min(Worksheets(i).Range("L2:L" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
    ActiveWorkbook.Worksheets(i).Cells(4, 17).Value = WorksheetFunction.Max(Worksheets(i).Range("J2:J" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
'Formating increase/decrease and greatest volumn table
    ActiveWorkbook.Worksheets(i).Cells(2, 17).Style = "Percent"
    ActiveWorkbook.Worksheets(i).Cells(3, 17).Style = "Percent"
'Finding ticker for increase/decrease and greatest volumn table
    For m = 2 To ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count
        If ActiveWorkbook.Worksheets(i).Cells(m, 12).Value = ActiveWorkbook.Worksheets(i).Cells(2, 17).Value Then
            ActiveWorkbook.Worksheets(i).Cells(2, 16).Value = ActiveWorkbook.Worksheets(i).Cells(m, 9).Value
        ElseIf ActiveWorkbook.Worksheets(i).Cells(m, 12).Value = ActiveWorkbook.Worksheets(i).Cells(3, 17).Value Then
            ActiveWorkbook.Worksheets(i).Cells(3, 16).Value = ActiveWorkbook.Worksheets(i).Cells(m, 9).Value
        ElseIf ActiveWorkbook.Worksheets(i).Cells(m, 10).Value = ActiveWorkbook.Worksheets(i).Cells(4, 17).Value Then
            ActiveWorkbook.Worksheets(i).Cells(4, 16).Value = ActiveWorkbook.Worksheets(i).Cells(m, 9).Value
        End If
    Next m
    
    k = 2
Next i

End Sub


