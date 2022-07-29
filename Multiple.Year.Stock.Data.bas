Attribute VB_Name = "Module1"
Sub MultipleYear()

Dim header() As Variant
Dim MainWs As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

header() = Array("Ticker ", "Date ", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", " ", " ", " ", "Ticker", "Value")

For Each MainWs In wb.Sheets
    With MainWs
    .Rows(1).Value = ""
    For i = LBound(header()) To UBound(header())
    .Cells(1, 1 + i).Value = header(i)
    
    Next i
    .Rows(1).Font.Bold = True
    .Rows(1).VerticalAlignment = xlCenter
    End With
    
Next MainWs

    For Each MainWs In Worksheets
    
        Dim Ticker As String
        Ticker = " "
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
        Dim Start_Price As Double
        Start_Price = 0
        Dim End_Price As Double
        End_Price = 0
        Dim Yearly_Price_Change As Double
        Yearly_Price_Change = 0
        Dim Yearly_Price_Percent As Double
        Yearly_Price_Percent = 0
        Dim Max_Ticker As String
        Max_Ticker = ""
        Dim Min_Ticker As String
        Min_Ticker = ""
        Dim Max_Percent As Double
        Max_Percent = 0
        Dim Min_Percent As Double
        Min_Percent = 0
        Dim Max_Volume_Ticker As String
        Max_Volume_Ticker = " "
        Dim Max_Volume As Double
        Max_Volume = 0
        Dim Min_Volume As Double
        Min_Volume = 0
      
        
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

Dim LastRow As Long

    LastRow = MainWs.Cells(Rows.Count, 1).End(xlUp).Row
    
    Start_Price = MainWs.Cells(2, 3).Value
    
For i = 2 To LastRow

    If MainWs.Cells(i + 1, 1).Value <> MainWs.Cells(i, 1).Value Then
    
        Ticker = MainWs.Cells(i, 1).Value
        
        End_Price = MainWs.Cells(i, 6).Value
        Yearly_Price_Change = End_Price - Start_Price
        
        If Start_Price <> 0 Then
            Yearly_Price_Percent = (Yearly_Price_Change / Start_Price) * 100
            
        End If
        
        Ticker_Volume = Ticker_Volume + MainWs.Cells(i, 7).Value
        MainWs.Range("I" & Summary_Table_Row).Value = Ticker
        MainWs.Range("J" & Summary_Table_Row).Value = Yearly_Price_Change
        
        If (Yearly_Price_Change > 0) Then
            MainWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
        ElseIf (Yearly_Price_Change <= 0) Then
            MainWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
        End If
        
        MainWs.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Price_Percent) & "%")
        
        MainWs.Range("L" & Summary_Table_Row).Value = Ticker_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Start_Price = MainWs.Cells(i + 1, 3).Value
        
        If (Yearly_Price_Percent > Max_Percent) Then
            Max_Percent = Yearly_Price_Percent
            Max_Ticker = Ticker
            
        ElseIf (Yearly_Price_Percent < Min_Percent) Then
            Min_Percent = Yearly_Price_Percent
            Min_Ticker = Ticker
            
        End If
        
        If (Ticker_Volume > Max_Volume) Then
            Max_Volume = Ticker_Volume
            Max_Volume_Ticker = Ticker
            
        End If
        
        Yearly_Price_Percent = 0
        Ticker_Volume = 0
        
    Else
        
        Ticker_Volume = Ticker_Volume + MainWs.Cells(i, 7).Value
        
    End If
    
Next i

    MainWs.Range("Q2").Value = CStr(Max_Percent) & "%"
    MainWs.Range("Q3").Value = CStr(Min_Percent) & "%"
    MainWs.Range("Q4").Value = Max_Volume
    MainWs.Range("P2").Value = Max_Ticker
    MainWs.Range("P3").Value = Min_Ticker
    MainWs.Range("P4").Value = Max_Volume_Ticker
    MainWs.Range("O2").Value = "Greatest % Increase"
    MainWs.Range("O3").Value = "Greatest % Decrease"
    MainWs.Range("O4").Value = "Greatest Total Volume"
    
Next MainWs
          
End Sub

