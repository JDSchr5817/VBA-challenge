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
Dim ticker_save_1 As String
Dim ticker_save_2 As String
Dim ticker_save_3 As String
Dim ticker_2 As String
Dim market_open As Double
Dim market_close As Double
Dim volume As Double
Dim volume_sum As Double
Dim save_1 As Double
Dim save_2 As Double
Dim max_vol As Double
Dim increase_1 As Double
Dim increase_2 As Double
Dim increase_save As Double
Dim decrease_1 As Double
Dim decrease_2 As Double
Dim decrease_save As Double
Dim volume_1 As Double
Dim volume_2 As Double
Dim ws As Worksheet
