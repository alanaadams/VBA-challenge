Attribute VB_Name = "Module1"
'This worksheet looping script was found on https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
'copyright 2009-2023 by www.extendoffice.com
Sub LoopWorkbook()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call alphabet_testing
    Next
    Application.ScreenUpdating = True
End Sub
        Sub alphabet_testing()

            'add new column labels
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "YearlyChange"
            Cells(1, 11).Value = "PercentChange"
            Cells(1, 12).Value = "Total Stock Volume"
    
            'count the number of rows
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
            'Summary Data Labels
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
    
            'Starting Value of Summary Data
            Cells(2, 17).Value = 0
            Cells(3, 17).Value = 0
            Cells(4, 17).Value = 0
    
    
    
            'Loop through all rows
            For I = 2 To lastrow
    
                'Add the Ticker info to the new colum
                Cells(I, 9).Value = Cells(I, 1).Value
                'Calucluate the YearlyChange and add to the new column
                Cells(I, 10).Value = Cells(I, 6).Value - Cells(I, 3).Value
                'Calculate the PercentChange and add to the new column
                Cells(I, 11).Value = (Cells(I, 10).Value / Cells(I, 3).Value)
                'Add the TotalStockVolume info to the new column
                Cells(I, 12).Value = Cells(I, 7).Value
        
                'Determine if cell has greatest percent increase
                If Cells(I, 11).Value > Cells(2, 17).Value Then
                    Cells(2, 17).Value = Cells(I, 11).Value
        
                'Determine if cell has greatest percent decrease
                ElseIf Cells(I, 11).Value < Cells(3, 17).Value Then
                    Cells(3, 17).Value = Cells(I, 11).Value
                Else
                End If
        
                'Determine if cell has greatest total volume
                If Cells(I, 12).Value > Cells(4, 17).Value Then
                    Cells(4, 17).Value = Cells(I, 12).Value
                Else
                End If
                
                'Add the ticker for the greatest percent increase
                If Cells(I, 11).Value = Cells(2, 17).Value Then
                    Cells(2, 16).Value = Cells(I, 1).Value
                Else
                End If
                
                'Add the ticker for the greatest percent decrease
                If Cells(I, 11).Value = Cells(3, 17).Value Then
                    Cells(3, 16).Value = Cells(I, 1).Value
                Else
                End If
                
                'Add the ticker for the greatest total volume
                If Cells(I, 12).Value = Cells(4, 17).Value Then
                    Cells(4, 16).Value = Cells(I, 1).Value
                Else
                End If
        
        
                'Formatting
                'PercentChange Cells in percent format
                Cells(I, 11).NumberFormat = "0.00%"
        
                ' Set the negative YearlyChange Cell Colors to Red
                If Cells(I, 10).Value < 0 Then
                    Cells(I, 10).Interior.ColorIndex = 3
        
                'Set the remaining YearlyChange Cell Colors to Green
                Else: Cells(I, 10).Interior.ColorIndex = 4
                End If
                       
            Next I
    
            'Greatest % Increase in percent format
            Cells(2, 17).NumberFormat = "0.00%"
            'Greatest % Decrease in percent format
            Cells(3, 17).NumberFormat = "0.00%"
            'Right align
            Range("O1:O4").HorizontalAlignment = xlRight

            End Sub

