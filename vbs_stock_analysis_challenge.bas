Attribute VB_Name = "Module1"
Sub StockAnalysis()


'Define variables

Dim Ticker As String
Dim year_open As Double
Dim year_close As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double
Dim start_data As Integer

Dim ws As Worksheet



For Each ws In Worksheets


    'Assign column headers for first analysis, format


    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").HorizontalAlignment = xlCenter
    ws.Range("I1:L1").Font.Bold = True


    'Assign integers for loop to starts

    start_data = 2
    previous_i = 1
    Total_Stock_Volume = 0


    'Find end row count

    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row


        'Analysis for each ticker

        For i = 2 To EndRow

            'Find last row for Ticker

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 

            'Add to table

            Ticker = ws.Cells(i, 1).Value
 

            'Continue

            previous_i = previous_i + 1

 
            'Get opening and closing values

            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value

            'Sum volume

            For j = previous_i To i

                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j


            'Conditional statement to calculate changes
 

            If year_open = 0 Then
                Percent_Change = year_close

            Else

                Yearly_Change = year_close - year_open

                Percent_Change = Yearly_Change / year_open

            End If

 
            'Place values in summary table

            ws.Cells(start_data, 9).Value = Ticker
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change
            
         'Format cell values

            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = Total_Stock_Volume

            
            'In the data summery when the first row task completed go to the next row

            start_data = start_data + 1

            'Reset variables

            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
 
            'Move i number to variable previous_i

            previous_i = i
 

        End If
 

    Next i


'Greatest summary table

    'Find last row

     kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

 
    'Define variables for greatest summary table


    Increase = 0
    Decrease = 0
    Greatest = 0


        'find max/min for percentage change and the max volume Loop

        For k = 3 To kEndRow
 
            'Define previous increment to check

            last_k = k - 1
 

            'Define current row for percentage

            current_k = ws.Cells(k, 11).Value
 

            'Define Previous row for percentage

            prevous_k = ws.Cells(last_k, 11).Value
 

            'greatest total volume row

            volume = ws.Cells(k, 12).Value
 

            'Prevous greatest volume row

            prevous_vol = ws.Cells(last_k, 12).Value


            'Find the increase

            If Increase > current_k And Increase > prevous_k Then
                Increase = Increase

            ElseIf current_k > Increase And current_k > prevous_k Then
                Increase = current_k
                increase_name = ws.Cells(k, 9).Value

            ElseIf prevous_k > Increase And prevous_k > current_k Then
                Increase = prevous_k
                increase_name = ws.Cells(last_k, 9).Value

            End If

 
            'Find the decrease

            If Decrease < current_k And Decrease < prevous_k Then
                Decrease = Decrease

            ElseIf current_k < Increase And current_k < prevous_k Then
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value

            ElseIf prevous_k < Increase And prevous_k < current_k Then
                Decrease = prevous_k
                decrease_name = ws.Cells(last_k, 9).Value

            End If



           'Find the greatest volume

            If Greatest > volume And Greatest > prevous_vol Then
                Greatest = Greatest

            ElseIf volume > Greatest And volume > prevous_vol Then
                Greatest = volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
                Greatest = prevous_vol
                greatest_name = ws.Cells(last_k, 9).Value

            End If


        Next k

 

    'Add column and row headers to greatest summary

    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"
    ws.Range("N1:P1").HorizontalAlignment = xlCenter
    

    'Input values

    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest


    'Format
    
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("P4").NumberFormat = "0"
    ws.Range("N1:P1").Font.Bold = True
    ws.Range("N2:N4").Font.Bold = True


'Conditional formatting for column colors
'End row calc

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

        For j = 2 To jEndRow

            'if greater than or less than zero

            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3

            End If


        Next j
 

ws.Columns("A:Z").AutoFit

'Move to next worksheet

Next ws


End Sub
