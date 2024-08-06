Attribute VB_Name = "Module1"
Sub stockData()

'This for loop will allow us to iterate over each worksheet in the workbook
For Each ws In Worksheets

    Dim WorksheetName As String
    'Assign the name of the current worksheet (ws) to the variable Worksheet 2018, 2019,2020
    WorksheetName = ws.Name
    
    'initialize all the variables needed for this exercise
    Dim i As Long
    Dim j As Long
    Dim Tcount As Long
    Dim Percent_Change As Double
    'Dim Greatest_Incr As Double
    'Dim Greatest_Decr As Double
    'Dim GreatestVol As Double
    
    
    'Set the value of a cell in the current worksheet ws to the string Ticker
    ws.Cells(1, 9).Value = "Ticker"
    'Set the value of a cell in the current worksheet ws to the string "Yearly Change"
    ws.Cells(1, 10).Value = "Yearly Change"
    'Set the value of a cell in the current worksheet ws to the string "Percent Change"
    ws.Cells(1, 11).Value = "Percent Change"
    'Set the value of a cell in the current worksheet ws to the string "Total Stock Volume"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'Keeps track of the current ticker
    Tcount = 2
    j = 2
    
    Dim LastRowA As Long
    
    'sets the variable LastRowA to the row number of the last non-empty cell in column A
    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'This is a loop that iterates through rows in the range from the 2nd row (i = 2) to the last row of data in column A
    For i = 2 To LastRowA
        
        'This condition will compare the value in the cell below (i +1 ) from the value i
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'If the previous line was True, this line will copy a value from one colmun to another column in a different row
            ws.Cells(Tcount, 9).Value = ws.Cells(i, 1).Value
            'This line will calculate the yearly change
            ws.Cells(Tcount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
              'If the cell value is less than zero then it will fill in the cell with red
              If ws.Cells(Tcount, 10).Value < 0 Then
                    ws.Cells(Tcount, 10).Interior.ColorIndex = 3
        
                Else
        
                    'Otherwise fill the cell background color to green
                    ws.Cells(Tcount, 10).Interior.ColorIndex = 4
                
                End If
            
            'Calculate and write percent change in column
            If ws.Cells(j, 3).Value <> 0 Then
                    Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
        
                    'These line will do Percent formating
                    ws.Cells(Tcount, 11).NumberFormat = "0.00%"
                    ws.Cells(Tcount, 11).Value = Percent_Change
                    
            Else
                        'If ws.Cells(j, 3).Value = 0 Then
                        ws.Cells(Tcount, 11).NumberFormat = "0.00%"
                        ws.Cells(Tcount, 11).Value = 0
                        
            End If
        
            'Calculate and write total volume in column
            ws.Cells(Tcount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            'Increase TickCount by 1
            Tcount = Tcount + 1
            
            'Set new start row of the ticker block
            j = i + 1
    
        End If
        
Next i

     Dim LastRowI As Long
     'Find last non-blank cell in column I
     LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
     Dim GreatestVol As Variant
     Dim Greatest_Incr As Variant
     Dim Greatest_Decr As Variant
     
     'This will give us the value of    GreatestVol, Greatest_Incr, Greatest_Decr
     GreatestVol = ws.Cells(2, 12).Value
     Greatest_Incr = ws.Cells(2, 11).Value
     Greatest_Decr = ws.Cells(2, 11).Value
    
    '
     For i = 2 To LastRowI
            If ws.Cells(i, 12).Value > GreatestVol Then
                    GreatestVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
            Else
                GreatestVol = GreatestVol
            
            End If
            
            
            If ws.Cells(i, 11).Value > Greatest_Incr Then
                Greatest_Incr = ws.Cells(i, 11).Value
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
            
            Else
            
                Greatest_Incr = Greatest_Incr
            
            End If
            
            If ws.Cells(i, 11).Value < Greatest_Decr Then
                Greatest_Decr = ws.Cells(i, 11).Value
                 ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
            
            Else
            
                 Greatest_Decr = Greatest_Decr
            
            End If
            
                ws.Cells(2, 16).Value = Format(Greatest_Incr, "Percent")
                ws.Cells(3, 16).Value = Format(Greatest_Decr, "Percent")
                ws.Cells(4, 16).Value = Format(GreatestVol, "Scientific")
                
        Next i

    Worksheets(WorksheetName).Columns("A:Z").AutoFit
Next ws

End Sub
