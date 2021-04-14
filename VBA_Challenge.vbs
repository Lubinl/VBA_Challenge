Attribute VB_Name = "Module1"
Sub Stock_Data()
'Create labels for columns to be added
Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("l1").Value = "Total Stock Volume"
'----------------------------------

'Define variables
Dim Ticker As String
Dim Annual_Opening_Price As Double
Dim Annual_Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Value As Double
Dim i As Long
Dim j As Long
Dim Ws As Worksheet
Dim LastRow As Long

i = 2
LastRow = Ws.Cells(Rows.count, "A").End(xlUp).Row


'Create Loop to serach rows for Ticker




End Sub
