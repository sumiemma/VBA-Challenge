Attribute VB_Name = "Module1"
Sub StockAnalysis()

Dim ws As Worksheet

Dim ticker As String

Dim Yearly_Change As Double
Yearly_Change = 0

Dim Percentage_Change As Double
Percentage_Change = 0

Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

For Each ws In Worksheets

Summary_Table_Row = 2

    For i = 2 To 759001

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
        
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
            Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(i - 250, 3).Value
        
            Percentage_Change = Yearly_Change / ws.Cells(i - 250, 3).Value
        
            ws.Range("K" & Summary_Table_Row).Value = ticker
        
            ws.Range("L" & Summary_Table_Row).Value = Yearly_Change
        
            ws.Range("M" & Summary_Table_Row).Value = Percentage_Change
        
            ws.Range("N" & Summary_Table_Row).Value = Total_Stock_Volume
        
            Summary_Table_Row = Summary_Table_Row + 1
        
            Total_Stock_Volume = 0
        
            Yearly_Change = 0
        
            Percentage_Change = 0
        
        
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        

        End If
    
    Next i
    
    i = 0

Next ws

For Each ws In Worksheets


    For i = 2 To 3001

        If ws.Cells(i, 12) >= 0 Then
            ws.Cells(i, 12).Interior.Color = vbGreen
    
        Else
            ws.Cells(i, 12).Interior.Color = vbRed
    
        End If

    Next i
    
    i = 0

Next ws


    



End Sub
Sub test()

Dim ticker As String

Dim Yearly_Change As Double
Yearly_Change = 0

Dim Percentage_Change As Double
Percentage_Change = 0

Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

For i = 2 To 753001

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        Yearly_Change = Cells(i, 6).Value - Cells(i - 250, 3).Value
        
        Percentage_Change = Yearly_Change / Cells(i - 250, 3).Value
        
        Range("K" & Summary_Table_Row).Value = ticker
        
        Range("L" & Summary_Table_Row).Value = Yearly_Change
        
        Range("M" & Summary_Table_Row).Value = Percentage_Change
        
        Range("N" & Summary_Table_Row).Value = Total_Stock_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Total_Stock_Volume = 0
        
        Yearly_Change = 0
        
        Percentage_Change = 0
        
        
    Else
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        

    End If
    
Next i



For i = 2 To 3001

    If Cells(i, 12) >= 0 Then
        Cells(i, 12).Interior.Color = vbGreen
    
    Else
        Cells(i, 12).Interior.Color = vbRed
    
    End If

Next i




'Total_Stock_Volume = Cells(i, 7).Value

'Range("N" & Summary_Table_Row).Value = WorksheetFunction.SumIfs(Range("G2:G503"), Range("A2:A503"), Cells(i, 1))



'Total_Stock_Volume = Range("N" & Summary_Table_Row).Value

'Next i

'Summary_Table_Row = Summary_Table_Row + 1


End Sub
