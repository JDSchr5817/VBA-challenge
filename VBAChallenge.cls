VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickers()

Dim ticker_1 As String
Dim ticker_2 As String
Dim ticker_save_1 As String
Dim ticker_save_2 As String
Dim ticker_save_3 As String
Dim market_open As Double
Dim market_close As Double
Dim vol As Double
Dim vol_sum As Double
Dim max_vol As Double
Dim vol_1 As Double
Dim vol_2 As Double
Dim save_1 As Double
Dim save_2 As Double
Dim increase_1 As Double
Dim increase_2 As Double
Dim increase_save As Double
Dim decrease_1 As Double
Dim decrease_2 As Double
Dim decrease_save As Double
Dim ws As Worksheet

For Each ws In Sheets
    
    ws.Cells(1, 9) = "ticker"
    ws.Cells(1, 10) = "yearly change"
    ws.Cells(1, 11) = "percent change"
    ws.Cells(1, 12) = "total volume"
    ws.Cells(1, 16) = "ticker"
    ws.Cells(1, 17) = "value"
    ws.Cells(2, 15) = "greatest percent increase"
    ws.Cells(3, 15) = "greatest percent decrease"
    ws.Cells(4, 15) = "greates total volume"
   
   For n = 1000000 To 2 Step -1
        If ws.Cells(n, 3) = "0" Then
        ws.Rows(n).EntireRow.Delete
        End If
    Next n
         
    save_1 = 2
    
    For j = 2 To 5000
   
    market_open = ws.Cells(save_1, 3).Value
    vol_sum = 0
    
        For i = save_1 To save_1 + 365
            ticker_1 = ws.Cells(i, 1)
            ticker_2 = ws.Cells(i + 1, 1)
            vol = ws.Cells(i, 7).Value
            vol_sum = vol_sum + vol

            If ticker_1 <> ticker_2 Then
            market_close = ws.Cells(i, 6).Value
    
            ws.Cells(j, 9).Value = ticker_1
            ws.Cells(j, 10).Value = market_close - market_open
            ws.Cells(j, 11).Value = ((market_close - market_open) / market_open)
            
            ws.Cells(j, 12).Value = vol_sum
        
            ticker_1 = ws.Cells(i + 1, 1).Value
            market_open = ws.Cells(i + 1, 3).Value
            save_1 = i + 1
            Exit For
            End If

    Next i
   
   If ws.Cells(j, 11).Value > 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 4
    ElseIf ws.Cells(j, 11).Value = 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 2
    Else
        ws.Cells(j, 11).Interior.ColorIndex = 3
    End If

Next j

max_vol = 0
For m = 2 To 4000
    
    vol_1 = ws.Cells(m, 11).Value
    vol_2 = ws.Cells(m + 1, 12).Value
    ticker_2 = ws.Cells(m + 1, 9).Value
        
    If vol_2 > max_vol Then
        max_vol = vol_2
        ticker_save_3 = ticker_2
    End If
             
Next m

increase_save = 0
For k = 2 To 4000
    
   ticker_1 = ws.Cells(k, 9).Value
   increase_1 = ws.Cells(k, 11).Value
   
        
    If increase_1 > increase_save Then
        increase_save = increase_1
        ticker_save_1 = ticker_1
    End If
    
Next k

decrease_save = 0
For l = 2 To 4000
    
    ticker_2 = ws.Cells(l, 9).Value
    decrease_1 = ws.Cells(l, 11).Value
    
        
    If decrease_1 < decrease_save Then
        decrease_save = decrease_1
        ticker_save_2 = ticker_2
    End If
    
               
Next l

    ws.Cells(2, 16).Value = ticker_save_1
    ws.Cells(2, 17).Value = increase_save
    ws.Cells(3, 16).Value = ticker_save_2
    ws.Cells(3, 17).Value = decrease_save
    ws.Cells(4, 16).Value = ticker_save_3
    ws.Cells(4, 17).Value = max_vol
    
Next ws


     
     
     
End Sub
