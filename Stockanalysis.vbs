Sub stockanalysis()

Dim ticker As String
Dim startdate As String
Dim enddate As String
Dim OpenValue As Double
Dim EndValue As Double
Dim Volume As Double
Dim percent_change As Double
Dim Summary_table_row As Integer
Summary_table_row = 2

Dim Greatest_total_volume As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double



Volume = 0


Dim lastrow As Long

'Loop through all sheets

For Each ws In ActiveWorkbook.Worksheets

  
    Dim WorksheetName As String
    WorksheetName = ActiveSheet.Name

    'Adding header for the summary table columns on each sheet

    ws.Range("I1").value = "Ticker"
    ws.Range("J1").value = "Quarterly Change"
    ws.Range("K1").value = "Percentage Change"
    ws.Range("L1").value = "Total Stock Volume"
    ws.Range("P1").value = "Ticker"
    ws.Range("Q1").value = "Value"

'Each Quarter's startdate and enddate

    If (WorksheetName = "Q1") Then
        startdate = "1/2/2022"
        enddate = "3/31/2022"

    ElseIf (WorksheetName = "Q2") Then

        startdate = "4/1/2022"
        enddate = "6/30/2022"

    ElseIf (WorksheetName = "Q3") Then
        startdate = "7/1/2022"
        enddate = "9/30/2022"

    Else

        startdate = "10/1/2022"
        enddate = "12/31/2022"

    End If
    
    'counting the lastrow
    lastrow = Cells(Rows.Count, 1).End(xlUp).row

    'Loop through rows for ticker and total volume value retrieval

    For i = 2 To lastrow

        If Cells(i + 1, 1).value <> Cells(i, 1).value Then

            ticker = Cells(i, 1).value

            Volume = Volume + Cells(i, 7)

            Range("I" & Summary_table_row).value = ticker

            Range("L" & Summary_table_row).value = Volume

            Summary_table_row = Summary_table_row + 1

            Volume = 0

        Else

            Volume = Volume + Cells(i, 7).value

        End If

    Next i

'calculate the difference between openvalue and closevalue and percentage change values

For r = 2 To lastrow   'for iterating on each row

    For j = 2 To 6     'for iterating on each column

    If (startdate = Cells(r, 2).value) Then  'taking the startdate of quarter and open value for each Ticker

        OpenValue = Cells(r, 3).value
        ticker = Cells(r, 1).value

        Exit For

    Else                'Continue the loop through rows and columns until the enddate of quarter is found for each Ticker and then take the closing value

        Dim QuarterEndDate As String
        Dim Quarterly_Change As Double

        QuarterEndDate = Cells(r, 2).value


        If (enddate = QuarterEndDate) Then    'once the open_value and close_value are found then perform the data filling of the summary table columns

            EndValue = Cells(r, 6).value

            Quarterly_Change = EndValue - OpenValue

            percent_change = (Quarterly_Change / OpenValue) * 100


            For c = 2 To 1501

                If (Range("I" & c).value = ticker) Then

                    Range("J" & c).value = Quarterly_Change


                    Range("K" & c).value = percent_change & "%"


                    If (Quarterly_Change > 0) Then

                       Range("J" & c).Interior.ColorIndex = 4

                       ElseIf (Quarterly_Change < 0) Then

                       Range("J" & c).Interior.ColorIndex = 3

                       End If

                    Exit For

                End If

            Next c

    End If

End If

Next j

Next r

Call max

Next ws


End Sub

Sub max()
 
 greatest_increase = WorksheetFunction.max(Range("K1:K1501"))
 
 
 Cells(2, 15).value = "Greatest % Increase"

 Cells(2, 17).value = greatest_increase
 
 greatest_decrease = WorksheetFunction.Min(Range("K1:K1501"))
 
  Cells(3, 15).value = "Greatest % Decrease"
  
  Cells(3, 17).value = greatest_decrease
 
  
  Cells(4, 15).value = "Greatest Total Volume "
  
 Greatest_total_volume = WorksheetFunction.max(Range("L2:L1501"))
  
 End Sub
