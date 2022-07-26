Attribute VB_Name = "Module1"
Sub alphabetical_testing():

Dim MainWs As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

Dim ticker_name As String

For Each MainWs In Worksheets

Dim total_stockVol As Double
    total_stockVol = 0
Dim Beg_Price As Double
    Beg_Price = 0
Dim End_Price As Double
    End_Price = 0
Dim Yearly_Price_Change As Double
    Yearly_Price_Change = 0
Dim Yearly_Price_Change_Percent As Double
    Yearly_Price_Change_Percent = 0




Dim OutputChartRow As Long
    OutputChartRow = 2
    
Dim LastRow As Long

    LastRow = MainWs.Cells(MainWs.Rows.Count, "A").End(xlUp).Row

Beg_Price = MainWs.Cells(2, 3).Value


For i = 2 To LastRow

        If MainWs.Cells(i + 1, 1).Value <> MainWs.Cells(i, 1).Value Then

 ticker_name = Cells(i, 1).Value
 End_Price = MainWs.Cells(i, 6).Value
 Yearly_Price_Change = End_Price - Beg_Price
 
            If Beg_Price <> 0 Then
    Yearly_Price_Change_Percent = (Yearly_Price_Change / Beg_Price) * 100
    
    End If
    
total_stockVol = total_stockVol + MainWs.Cells(i, 7).Value

MainWs.Range("I" & OutputChartRow).Value = ticker_name

MainWs.Range("J" & OutputChartRow).Value = Yearly_Price_Change

If (Yearly_Price_Change > 0) Then
MainWs.Range("J" & OutputChartRow).Interior.ColorIndex = 4

ElseIf (Yearly_Price_Change <= 0) Then
MainWs.Range("J" & OutputChartRow).Interior.ColorIndex = 3
End If

MainWs.Range("K" & OutputChartRow).Value = (CStr(Yearly_Price_Change_Percent) & "%")

MainWs.Range("L" & OutputChartRow).Value = total_stockVol

OutputChartRow = OutputChartRow + 1

Beg_Price = MainWs.Cells(i + 1, 3).Value
   

Else

total_stockVol = total_stockVol + MainWs.Cells(i, 7).Value

End If

Next i

MainWs.Cells(1, 9).Value = "ticker symbol"
MainWs.Cells(1, 10).Value = "yearly change($)"
MainWs.Cells(1, 11).Value = "percent change"
MainWs.Cells(1, 12).Value = "total stock volume"
MainWs.Columns("I:L").AutoFit

Next MainWs


End Sub




