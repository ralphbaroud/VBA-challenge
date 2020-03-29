Sub The_VBA_of_Wall_Street()
'NOTE: Some terms from the README were changed to the ones used in the financial industry. Ultimately, the result is the same.
'Applying the below for all worksheets.
For Each ws In Worksheets
'Defining needed variables.
    Dim Ticker As String
    Dim i As Double
    Dim j As Double
    Dim Trading_Volume As Double
    Dim CRNCY_Change As Double
    Dim Percentage As Double
    Dim Row_Summary_Table As Double
    Dim Last_Row_Value As Double
    Dim Previous As Double
    Dim Largest_Increase As Double
    Dim Largest_Decrease As Double
    Dim Largest_Trading_Volume As Double
    Dim MAX As Double
    Dim MIN As Double
    Dim VOL As Double
'Adding the desired labels as per the provided images with some formatting for aesthetic purposes.
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly (Currency) Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Trading Volume"
    ws.Range("I1:L1").Font.Bold = True
    ws.Range("I1:L1").VerticalAlignment = xlCenter
    ws.Range("I1:L1").HorizontalAlignment = xlCenter
    ws.Range("O2").Value = "Largest Percentage Increase"
    ws.Range("O3").Value = "Largest Percentage Decrease"
    ws.Range("O4").Value = "Largest Trading Volume"
    ws.Range("O2:O4").Font.Bold = True
    ws.Range("O2:O4").VerticalAlignment = xlCenter
    ws.Range("O2:O4").HorizontalAlignment = xlLeft
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P1:Q1").Font.Bold = True
    ws.Range("P1:Q1").VerticalAlignment = xlCenter
    ws.Range("P1:Q1").HorizontalAlignment = xlCenter
'Obtaining the value for the very last row in each worksheet and setting all needed variables to respective constants.
    Last_Row_Value = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Trading_Volume = 0
    Largest_Increase = 0
    Largest_Decrease = 0
    Largest_Trading_Volume = 0
    Previous = 2
    Row_Summary_Table = 2
'Beginning of loop.
        For i = 2 To Last_Row_Value
            Trading_Volume = Trading_Volume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Row_Summary_Table).Value = Ticker
                ws.Range("L" & Row_Summary_Table).Value = Trading_Volume
                ws.Range("L" & Row_Summary_Table).NumberFormat = "###,###,###,##0"
                ws.Range("L" & Row_Summary_Table).HorizontalAlignment = xlCenter
                ws.Range("L" & Row_Summary_Table).VerticalAlignment = xlCenter
                Trading_Volume = 0

                    If ws.Range("C" & Previous) >= ws.Range("F" & i) Then

                        CRNCY_Change = -(ws.Range("C" & Previous) - ws.Range("F" & i))
                        ws.Range("J" & Row_Summary_Table).Value = CRNCY_Change
                        ws.Range("J" & Row_Summary_Table).HorizontalAlignment = xlCenter
                        ws.Range("J" & Row_Summary_Table).VerticalAlignment = xlCenter

                    Else

                        CRNCY_Change = (ws.Range("F" & i) - ws.Range("C" & Previous))
                        ws.Range("J" & Row_Summary_Table).Value = CRNCY_Change
                        ws.Range("J" & Row_Summary_Table).HorizontalAlignment = xlCenter
                        ws.Range("J" & Row_Summary_Table).VerticalAlignment = xlCenter

                    End If

                    If CRNCY_Change >= 0 Then

                        ws.Range("J" & Row_Summary_Table).Interior.ColorIndex = 4

                    Else
                        ws.Range("J" & Row_Summary_Table).Interior.ColorIndex = 3

                    End If

                    If CRNCY_Change = 0 Or ws.Range("C" & Previous) = 0 Then
                        
                        Percentage = "0"
                    Else

                        Percentage = ws.Range("J" & Row_Summary_Table) / ws.Range("C" & Previous)
                        ws.Range("K" & Row_Summary_Table).Value = Percentage
                        ws.Range("K" & Row_Summary_Table).NumberFormat = "0.00 %"
                        ws.Range("K" & Row_Summary_Table).HorizontalAlignment = xlCenter
                        ws.Range("K" & Row_Summary_Table).VerticalAlignment = xlCenter

                    End If
                
                Row_Summary_Table = Row_Summary_Table + 1
                Previous = i + 1
            
            End If
        
        Next i

    'Adding values and running a loop for the required largest increases and largest decrease cell.
    Last_Row_Value = ws.Cells(Rows.Count, 11).End(xlUp).Row

    Largest_Increase = MAX
    Largest_Decrease = MIN
    Largest_Trading_Volume = VOL
    MAX = ws.Range("K2")
    MIN = ws.Range("K2")
    VOL = ws.Range("L2")
    ws.Range("P2").Value = ws.Range("I2")
    ws.Range("P3").Value = ws.Range("I2")
    ws.Range("P4").Value = ws.Range("L2")

    For i = 2 To Last_Row_Value

        If ws.Range("K" & i).Value > MAX Then
        MAX = ws.Range("K" & i).Value
        ws.Range("Q2").Value = MAX
        ws.Range("P2").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("K" & i).Value < MIN Then
        MIN = ws.Range("K" & i).Value
        ws.Range("Q3").Value = MIN
        ws.Range("P3").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("L" & i).Value > VOL Then
        VOL = ws.Range("L" & i).Value
        ws.Range("Q4").Value = VOL
        ws.Range("P4").Value = ws.Range("I" & i).Value
        End If

    Next i

'Additional Formating
ws.Range("P2:Q5").HorizontalAlignment = xlCenter
ws.Range("P2:Q5").VerticalAlignment = xlCenter
ws.Range("Q2:Q3").NumberFormat = "0.00 %"
ws.Range("Q4").NumberFormat = "###,###,###,##0"

'Unintended loop I was forced to create as a last resort since the code was returning a void cell when there is no change in the price. Corrects the problem.
For j = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
    If ws.Range("J" & j) = 0 Then
    ws.Range("K" & j) = 0#
    ws.Range("K" & j).NumberFormat = "0.00 %"
    ws.Range("K" & j).HorizontalAlignment = xlCenter
    ws.Range("K" & j).VerticalAlignment = xlCenter
    End If
Next j

ws.Columns("A:Q").AutoFit
Next ws
End Sub