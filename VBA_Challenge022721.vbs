Sub stock_report()

    'loop through each Worksheet
    For Each ws in Worksheets

        ' --------------------------------------------
        ' INSERT THE STATE
        ' --------------------------------------------

        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        ' MsgBox WorksheetName

        ' Set an initial variable for holding the brand name
        Dim Ticker_Symbol As String

        ' Set an initial variable for holding yearly change in stock price
        Dim OpeningRowNumber As long
        OpeningRowNumber = 2
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Stock_Volume As Double
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim Greatest_Total_Volume As Double
        Dim Ticker_Greatest_Increase As String
        Dim Ticker_Greatest_Decrease As String
        Dim Ticker_Greatest_Total_Volume As String         
        Stock_Volume = 0
        GreatestIncrease  = 0
        GreatestDecrease = 0
        Greatest_Total_Volume = 0


        ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 1

        'Print summary table header
        ws.Range("I" & Summary_Table_Row).Value =  "Ticker"
        ws.Range("J" & Summary_Table_Row).Value =  "Yearly_Change"
        ws.Range("K" & Summary_Table_Row).Value =  "Percent_Change"
        ws.Range("L" & Summary_Table_Row).Value =  "Volume"
        ws.Range("O" & 2).Value =  "Greatest%Increase"
        ws.Range("O" & 3).Value =  "Greatest%Decrease"
        ws.Range("O" & 4).Value =  "GreatestTotalVolume"
        ws.Range("P" & 1).Value =  "Ticker"
        ws.Range("Q" & 1).Value =  "Value"

        Summary_Table_Row = Summary_Table_Row + 1

        ' Loop through all stock tikers
        For i = 2 To LastRow

            ' Check if the stock tiker are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                Ticker_Symbol = ws.Cells(i, 1).Value

                ' set yearly change value
                Opening_Price = ws.Cells(OpeningRowNumber,3).Value
                Closing_Price = ws.Cells(i,6).Value
                Yearly_Change = Closing_Price - Opening_Price

                ' Set Percent Change value
                If Opening_Price = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yearly_Change / Opening_Price
                End If
            
                ' Add to the stock volume
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value


                ' Print the ticker symbol in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol

                ' Print the Brand Amount to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                ' Print the Brand Amount to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = format(Percent_Change, "Percent")
                If Percent_Change > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf Percent_Change < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If


                ' Print the Brand Amount to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                ' Set opening row number
                OpeningRowNumber = i + 1
                
                ' Reset the Stock_Volume
                Stock_Volume = 0

                ' If the cell immediately following a row is the same brand...
            Else
                ' Add to the Stock Volume
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            End If

        Next i

        'Print summary table
        For i = 2 to Summary_Table_Row
            If  ws.Cells(i,"K").Value > GreatestIncrease Then 
                GreatestIncrease = ws.Cells(i,"K").Value
                Ticker_Greatest_Increase = ws.Cells(i,"I")               
            Else
                If ws.Cells(i,"K").Value < GreatestDecrease  Then
                    GreatestDecrease = ws.Cells(i,"K").Value
                    Ticker_Greatest_Decrease = ws.Cells(i,"I")
                End If
            End If

            If  ws.Cells(i,"L").Value > Greatest_Total_Volume Then 
                Greatest_Total_Volume = ws.Cells(i,"L").Value
                Ticker_Greatest_Total_Volume = ws.Cells(i,"I") 
            End If

        Next i

            'Print out the Ticker and values for summary table
            ws.Range("Q" & 2).Value =  format(GreatestIncrease,"Percent")
            ws.Range("Q" & 3).Value =  format(GreatestDecrease,"Percent")
            ws.Range("Q" & 4).Value =  Greatest_Total_Volume
            ws.Range("P" & 2).Value =  Ticker_Greatest_Increase
            ws.Range("P" & 3).Value =  Ticker_Greatest_Decrease
            ws.Range("P" & 4).Value =  Ticker_Greatest_Total_Volume
            ws.columns("I:Q").AutoFit
    Next ws

    MsgBox ("Analysis Report is Complete")

End Sub
