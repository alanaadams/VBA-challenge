# stock-analysis
Challenge 2 VBA Project

Instructions
Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock.

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

References
'This worksheet looping script was found on https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
'copyright 2009-2023 by www.extendoffice.com
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
    'your code here
End Sub