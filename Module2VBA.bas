Attribute VB_Name = "Module2"
Sub Stock_Data_Challenge()

'Outline Variables
'---

'Variables Used: (Ticker, Opening Price, Closing Price, Quarterly Change, Percentage Change, Total Stock Volume)
Dim ticker As String
Dim open_price As Double
Dim closing_price As Double
Dim qc As Double
Dim pc As Double
Dim tsv As Double

'Variables Used cont.: (Greatest Total Volume, Greatest Increase, Greatest Decrease)
Dim gtv As Double
Dim PreviousStockPrice As Long
Dim table_summary_row As Long
Dim greatest_increase As Double
Dim greatest_decrease As Double

Dim ws As Worksheet
For Each ws In Worksheets

'Column Titles
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'More Columns
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Determine the Quarterly Change, Percentage Change, and Total Stock Value for each stock group)
'---

'Values and Variables for first loop
tsv = 0
table_summary_row = 2
PreviousStockPrice = 2

'Input Value of the Last Row for Column A
EndRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through the rows of sheet
For u = 2 To EndRowA
    'Find Total Stock Volume
    tsv = tsv + ws.Cells(u, 7).Value
    
    'Formula to record for change in ticker with name and tsv. Then have tsv set back to zero
    If ws.Cells(u + 1, 1).Value <> ws.Cells(u, 1).Value Then
        ticker = ws.Cells(u, 1).Value
        ws.Range("I" & table_summary_row).Value = ticker
        ws.Range("L" & table_summary_row).Value = tsv
        tsv = 0
        
        'Include Opening Price, Closing Price, Quarterly Change, and Percentage Change
        open_price = ws.Range("C" & PreviousStockPrice)
        close_price = ws.Range("F" & u)
        qc = close_price - open_price
        ws.Range("J" & table_summary_row).Value = qc
        
        'Use an If Statement to find Percentage Change
        If open_price = 0 Then
            pc = 0
        Else
            open_price = ws.Range("C" & PreviousStockPrice)
            pc = qc / open_price
        End If
        
        'Format Percentage Change into summary table. Use "%"
        ws.Range("K" & table_summary_row).Value = pc
        ws.Range("K" & table_summary_row).NumberFormat = "0.00%"
        
        'Use If Statement for Conditional Formatting
        If ws.Range("J" & table_summary_row).Value >= 0 Then
        ws.Range("J" & table_summary_row).Interior.ColorIndex = 4
        Else
        ws.Range("J" & table_summary_row).Interior.ColorIndex = 3
        
        End If
        'Continue to next row til completion
        table_summary_row = table_summary_row + 1
        PreviousStockPrice = u + 1
        
    End If
    Next u
    
'Loop through to find Greatest Percentage Increase, Decrease, and Greatest Total Volume
'---
'Loop Values
greatest_increase = 0
greatest_decrease = 0
gtv = 0

'Now, Input Value of the Last Row for Column K
EndRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row

For u = 2 To EndRowK

    'Find Greatest Total Volume
    If ws.Range("L" & u).Value > gtv Then
       gtv = ws.Range("L" & u).Value
       ws.Range("Q4").Value = gtv
       ws.Range("P4").Value = ws.Range("I" & u).Value
       
    End If
    
    If ws.Range("K" & u).Value < greatest_decrease Then
        greatest_decrease = ws.Range("K" & u).Value
        ws.Range("Q3").Value = greatest_decrease
        ws.Range("P3").Value = ws.Range("I" & u).Value
        
    End If
    
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
Next u
'Continue Using Same Loops For the rest of the Worksheets
'---
Next ws
 
    
        
        
End Sub
