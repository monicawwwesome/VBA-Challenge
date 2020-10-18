Attribute VB_Name = "Module1"
Sub Yearly_Change_Testrun()
'Dim everything
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TickerTableRow As Integer
        TickerTableRow = 2
    Dim TotalVol As Double
        TotalVol = 0
    Dim YearlyChange As Double
       
'Determine last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Create for loop
For i = 2 To LastRow
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        ClosePrice = Cells(i, 6).Value
        
        For j = 2 To LastRow
        
        'If the ticker is new, store first open price
        If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
        
            OpenPrice = Cells(j, 3).Value
        
        'Calculate yearly change
        YearlyChange = OpenPrice - ClosePrice
        
        
        'Print yearly change
        Range("J" & TickerTableRow) = YearlyChange
        
        'New row in ticker table
        TickerTableRow = TickerTableRow + 1
        
        End If
        Next j
    End If
    Next i
    
    
End Sub
