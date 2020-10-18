Attribute VB_Name = "Module1"
Sub vba_hw()
'Dim everything
    Dim Ticker As String
    Dim LastRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim YearlyPercentage As Long
    Dim TickerTableRow As Double
        TickerTableRow = 2
    Dim TotalVol As Double
        TotalVol = 0
    Dim day As Double
        day = 0
       
'Determine last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Format ticker table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

     
'Create for loop
For i = 2 To LastRow
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        'Set ticker variable
        Ticker = Cells(i, 1).Value
        
        'Add vol to total vol
        TotalVol = TotalVol + Cells(i, 7).Value
        
        'Print if a different ticker found
        Range("I" & TickerTableRow).Value = Cells(i, 1).Value
        
        'Print total vol to the table
        Range("L" & TickerTableRow).Value = TotalVol
        
        'Record closing price
        ClosePrice = Cells(i, 6).Value
        
        'Calculate yearly change from opening price to closing price
        YearlyChange = ClosePrice - OpenPrice
        
        'Calculate Percentage change from opening price to closing price
        If OpenPrice <> 0 Then
        YearlyPercentage = YearlyChange / OpenPrice
        
        
        End If
         
        'Print yearly change
        Range("J" & TickerTableRow).Value = YearlyChange
        
        'Print yearly percentage
        Range("K" & TickerTableRow).Value = Round(YearlyPercentage, 2)
        Range("K" & TickerTableRow).NumberFormat = "0.00%"
        
        'New row in ticker table
        TickerTableRow = TickerTableRow + 1
        
        'Reset the total vol
        TotalVol = 0
        
        'Reset day of year
        day = 0
        

    Else
        
        'Add vol to total vol
        TotalVol = TotalVol + Cells(i, 7).Value
        
        'Check if it is the first day of the year
        day = day + 1
        
        If day = 1 Then
            OpenPrice = Cells(i, 3).Value
        End If
        
   End If
   
Next i

'Formating and coloring
For j = 2 To LastRow
    
    If Cells(j, 10).Value >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        
    Else
        Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
Next j
    
        

End Sub
