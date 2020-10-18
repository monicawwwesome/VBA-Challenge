Attribute VB_Name = "Module1"
Sub vba_hw()
'Dim everything
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TickerTableRow As Integer
        TickerTableRow = 2
    Dim TotalVol As Double
        TotalVol = 0
       
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
        
        'Ticker Column and Total Stock Volume Column
        '==============================================
        
        'Set ticker variable
        Ticker = Cells(i, 1).Value
        
        'Add vol to total vol
        TotalVol = TotalVol + Cells(i, 7).Value
        
        'Print if a different ticker found
        Range("I" & TickerTableRow).Value = Cells(i, 1).Value
        
        'Print total vol to the table
        Range("L" & TickerTableRow).Value = TotalVol
        
        'New row in ticker table
        TickerTableRow = TickerTableRow + 1
        
        'Reset the total vol
        TotalVol = 0
        
        'Yearly Change Column
        '==============================================
        

        
        
        'Yearly Percentage Column
        '==============================================
    Else
        
        'Add vol to total vol
        TotalVol = TotalVol + Cells(i, 7).Value
        
        
        
        
        
        
        

        


'Calculate yearly change from opening price to closing price

'Calculate Percentage change from opening price to closing price

'Conditional formatting color coding with yearly change

'End if
    End If
'Next i
Next i


End Sub
