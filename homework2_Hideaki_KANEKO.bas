Attribute VB_Name = "Module1"
Sub homework2()

Dim i As Long
Dim LastRow As Long 'Define Last Low

Dim initial As Double 'Opening Price of the Stock at the beginning of the Year
Dim last As Double 'Closing Price of the Stock at the end of the Year
Dim delta As Double 'Stock Price Change over the Year
Dim percent As Double 'Percent Change over the Year
Dim total As Double 'Total Stock Volume over the Year
Dim counter As Integer 'This "counter" is used for output

Dim max_percent As Double
Dim min_percent As Double
Dim max_total As Double
Dim max_percent_ticker As String
Dim min_percent_ticker As String
Dim max_total_ticker As String

'--------[Extra] Columns Label & Width Setting--------
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Columns(10).AutoFit
Columns(11).AutoFit
Columns(12).AutoFit
Columns(15).ColumnWidth = 19.5
'------------------------------------------------

'Initial value setting
initial = 0
last = 0
delta = 0
percent = 0
total = 0
counter = 2

max_percent = 0
min_percent = 0
max_total = 0

' Determine the Last Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row


'--------Main Loop--------
For i = 2 To LastRow

    'Caculate Total Stock Volume
    total = total + Cells(i, 7).Value
    
    'If it is in the first raw of the ticker, update Opening Price of the Stock
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        initial = Cells(i, 3).Value
    
    End If
        
    'If it reaches boundary of ticker change,
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        last = Cells(i, 6).Value 'Update Closing Price of the Stock
        delta = last - initial 'Calculate "delta"
        
        If initial = 0 Then 'This "If" section is for avoiding overflow due to "div by zero"
            percent = 0
        
        Else
            percent = delta / initial 'Calculate "percent"
        
        End If
                
        'Output Following Items,
        Cells(counter, 9).Value = Cells(i, 1).Value 'Ticker
        Cells(counter, 10).Value = delta 'Yearly Change
        Cells(counter, 11).Value = percent 'Percent Change
        Cells(counter, 12).Value = total 'Total Stock Volume
        
        '--------Extra Challenge Section 1--------
        'Update Max & Min "percent"
        If percent > max_percent Then
            max_percent = percent
            max_percent_ticker = Cells(i, 1).Value
            
        ElseIf percent < min_percent Then
            min_percent = percent
            min_percent_ticker = Cells(i, 1).Value
            
        End If
        
        'Update Max "total'
        If total > max_total Then
            max_total = total
            max_total_ticker = Cells(i, 1).Value
            
        End If
        '-----End of Extra Challenge Section 1-----
        
        'Format "Yearly Change" & "Percent Change" Column
        If delta > 0 Then
            Cells(counter, 10).Interior.ColorIndex = 4 'Format positive change in "Green"
            
        ElseIf delta < 0 Then
            Cells(counter, 10).Interior.ColorIndex = 3 'Format negative change in "Red"
            
        End If
                    
        Cells(counter, 11) = Format(Cells(counter, 11).Value, "0.00%") 'Change it to percent indication
                            
        'Reset Variables
        initial = 0
        last = 0
        delta = 0
        percent = 0
        total = 0
        
        'Proceed "counter"
        counter = counter + 1
        
    End If
    
Next i
'--------End of Main Loop--------


'--------Extra Challenge Section 2--------
'Output Following Items
Cells(2, 15).Value = "Greatest % Increase"
Cells(2, 16).Value = max_percent_ticker
Cells(2, 17).Value = max_percent
Cells(2, 17) = Format(Cells(2, 17).Value, "0.00%") 'Change it to percent indication

Cells(3, 15).Value = "Greatest % Decrease"
Cells(3, 16).Value = min_percent_ticker
Cells(3, 17).Value = min_percent
Cells(3, 17) = Format(Cells(3, 17).Value, "0.00%") 'Change it to percent indication

Cells(4, 15).Value = "Greatest Total Volume"
Cells(4, 16).Value = max_total_ticker
Cells(4, 17).Value = max_total
'-----End of Extra Challenge Section 2-----

End Sub
