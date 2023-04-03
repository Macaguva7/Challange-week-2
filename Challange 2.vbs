Sub WorksheetLoop()
    
    Const COLOR_GREEN As Integer = 4
    Const COLOR_RED As Integer = 3
    Const VOL_COL As Integer = 7
    ' Dim HEADERS() As String
    ' HEADERS = Split(",Ticker,Yearly_Change,Percent_Change,Stock_Volume,,,,Ticker,Value", ",")
    
    'WS
    Dim WS As Worksheet
    Dim wb As Workbook
    Dim header_index As Integer
    Dim header_column As Integer
    Dim Ticker_Name As String
    Dim Total_Ticker_Volume As Double
    Dim End_Price As Double
    Dim Yearly_Price_Change As Double
    Dim Yearly_Price_Percent As Double
    Dim Max_Ticker_Name As String
    Dim Max_Percent As Double
    Dim Min_Percent As Double
    Dim Max_Volume_Ticker_Name As String
    Dim Max_Volume As Double
    Dim Summary_Table_Row As Long
    Dim Lastrow As Long
    Set wb = ActiveWorkbook
    
'    For Each WS In wb.Sheets
'        With WS
'            .Rows(1).Value = ""
'            For i = LBound(headers()) To UBound(headers())
'                .Cells(1, 1 + i).Value = headers
'        Next i
'       .Rows(1).Font.Bold = True
'       .Rows(1).VerticalAlignment = xlCenter
'        End With
'       Next WS
'
    'loop through all of the workshhet
    For Each WS In Worksheets
        WS.Activate
        
        'Column Headers
        For header_index = 1 To UBound(HEADERS)
            header_column = VOL_COL + 1 + header_index
            Cells(1, VOL_COL + 1 + header_index).Value = HEADERS(header_index)
        Next header_index
        
        ' Set initial variables
        Ticker_Name = ""
        Beg_Price = 0
        End_Price = 0
        Yearly_Price_Change = 0
        Yearly_Price_Percent = 0
        Max_Ticker_Name = ""
        Min_Ticker_Name = ""
        Max_Percent = 0
        Min_Percent = 0
        Max_Volume_Ticker_Name = ""
        Max_Volume = 0
        
        'Row definition
        Summary_Table_Row = 2
        Lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To Lastrow
        
            'Ticker Name If
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                Ticker_Name = WS.Cells(i, 1).Value
                
                'Calculate End Price
                End_Price = WS.Cells(i, 6).Value
                Yearly_Price_Change = End_Price - Beg_Price
                
                If Beg_Price <> 0 Then
                    Yearly_Price_Change_Percent = (Yearly_Price_Change / Beg_Price) * 100
                End If
                
                Total_Ticker_Volume = Total_Ticker_Volume + WS.Cells(i, 7).Value
                WS.Range("I" & Summary_Table_Row).Value = Ticker_Name
                WS.Range("J" & Summary_Table_Row).Value = Yearly_Price_Change
                
                If (Yearly_Price_Change > 0) Then
                    WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = COLOR_GREEN
                ElseIf (Yearly_Price_Change <= 0) Then
                    WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = COLOR_RED
                End If
                
                 Total_Ticker_Volume = Total_Ticker_Volume + WS.Cells(i, 7).Value
                
                'Print Ticker in Summary Table, Column I
                WS.Range("I" & Summary_Table_Row).Value = Yearly_Price_Change
                
                'Print Yearly Price Change in Summary Table, column J
                WS.Range("J" & Summary_Table_Row).Value = Yearly_Price_Change
                
                If (Yearly_Price_Change > 0) Then
                    WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = COLOR_GREEN
                
                ElseIf (Yearly_Price_Change <= 0) Then
                    WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = COLOR_RED
                End If
                
                'Print Yearly Price Change as Percent in summary table, column K
                WS.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Price_Change_Percent) & "%")
                
                'Print Total Stock Volume in Summary Table, column L
                WS.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                'Add 1 to summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Get Beginning Price
                Beg_Price = WS.Cells(i, 3).Value
                
                'For Final Calculations
                
                If (Yearly_Price_Change_Percent > Max_Percent) Then
                    Max_Percent = Yearly_Price_Change_Percent
                    Max_Ticker_Name = Ticker_Name
                
                ElseIf (Yearly_Price_Change_Percent < Min_Percent) Then
                    Min_Percent = Yearly_Price_Change_Percent
                    Min_Ticker_Name = Ticker_Name
                    
                
                
                End If
                
                'Value Reset
                Yearly_Price_Change_Percent = 0
                Total_Ticker_Volume = 0
                
            ' Else if in the next ticker name, enter new ticker stock value
            Else
                Total_Ticker_Volume = Total_Ticker_Volume + WS.Cells(i, 7).Value
            End If
            
        Next i
        
        'Print Values in  Cells
        WS.Range("O2").Value = "Greatest%Increase"
        WS.Range("O3").Value = "Greatest%Decrease"
        WS.Range("O4").Value = "Greatest_Total_Volume"
        
        WS.Range("P2").Value = Max_Ticker_Name
        WS.Range("P3").Value = Min_Ticker_Name
        
        WS.Range("Q2").Value = (CStr(Max_Percent) & "%")
        WS.Range("Q3").Value = (CStr(Min_Percent) & "%")
        WS.Range("Q4").Value = Max_Volume
    
    Next WS
End Sub
