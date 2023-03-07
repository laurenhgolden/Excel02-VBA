Attribute VB_Name = "Module1"
Sub MultipleYearStock1():

' Name variables and assign values
Dim ws As Worksheet
Dim Ticker As String
Dim OpeningPrice As Double
OpeningPrice = (Cells(2, 3).Value)
Dim ClosingPrice As Double
Dim PriceChange As Double
Dim PercentageChange As Double
Dim TotalVolume As Double
TotalVolume = 0
Dim GreatestPercentageIncrease As Double
GreatestPercentageIncrease = 0
Dim GreatestIncreaseCompany As String
Dim GreatestPercentageDecrease As Double
GreatestPercentageDecrease = 0
Dim GreatestDecreaseCompany As String
Dim GreatestVolume As Double
GreatestVolume = 0
Dim GreatestVolumeCompany As String
Dim SummaryTableRow As Integer
SummaryTableRow = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row


' Create iteration for all ws in file
    For Each ws In ThisWorkbook.Worksheets
        'Assign Headings for each sheet under which systems will calculate
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
'Create iterations for each individual ws to calculate necessary values
        
        For i = 2 To lastrow
            'Conditional statement
            If (ws.Cells(i + 1, 1).Value) <> (ws.Cells(i, 1).Value) Then
                'Value changes/calculations
                Ticker = (ws.Cells(i, 1).Value)
                ClosingPrice = (ws.Cells(i, 6).Value)
                PriceChange = ClosingPrice - OpeningPrice
                PercentageChange = (PriceChange / OpeningPrice)
                    If PercentageChange > GreatestPercentageIncrease Then
                        GreatestPercentageIncrease = PercentageChange
                        GreatestIncreaseCompany = Ticker
                    End If
                    If PercentageChange < 0 Then
                        GreatestPercentageDecrease = PercentChange
                        GreatestDecreaseCompany = Ticker
                    End If
            
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    If TotalVolume > GreatestVolume Then
                        GreatestVolume = TotalVolume
                        GreatestVolumeCompany = Ticker
                    End If
                
                'Assign created values to Table
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("J" & SummaryTableRow).Value = PriceChange
                    If PriceChange > 0 Then
                        ws.Range("J" & SummaryTableRow).Interior.Color = vbGreen
                    Else: ws.Range("J" & SummaryTableRow).Interior.Color = vbRed
                    End If
                ws.Range("K" & SummaryTableRow).Value = PercentageChange
                    If PercentageChange > 0 Then
                        ws.Range("K" & SummaryTableRow).Interior.Color = vbGreen
                    Else: ws.Range("K" & SummaryTableRow).Interior.Color = vbRed
                    End If
                ws.Range("L" & SummaryTableRow).Value = TotalVolume
            
                'Reset for next ticker
                OpeningPrice = (Cells(i + 1, 3).Value)
                TotalVolume = 0
                SummaryTableRow = SummaryTableRow + 1
            
                Else
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
            End If
            ws.Range("P2").Value = GreatestIncreaseComany
            ws.Range("P3").Value = GreatestDecreaseCompany
            ws.Range("P4").Value = GreatestVolumeCompany
            ws.Range("Q2").Value = GreatestPercentageIncrease
            ws.Range("Q3").Value = GreatestPercentageDecrease
            ws.Range("Q4").Value = GreatestVolume
        Next i
        
        SummaryTableRow = 2
        OpeningPrice = (Cells(2, 3).Value)
        GreatestPercentageIncrease = 0
        GreatestPercentageDecrease = 0
        GreatestVolume = 0
        GreatestIncreaseCompany = ""
        GreatestDecreaseCompany = ""
        GreatestVolumeCompany = ""
        
    Next ws
            
End Sub
