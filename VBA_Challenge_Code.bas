Attribute VB_Name = "Module1"
Sub Wall_Street()

'Introduce headers for all workbooksheets

For Each ws In ThisWorkbook.Worksheets
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

'Define Variables
Dim Ticker As String
Dim Year_open As Double
Dim Year_close As Double
Dim Percent_change As Double
Dim Yearly_change As Double
Dim ct As Double
ct = 2

Dim Total_vol As Double
Total_vol = 0
Dim Summary_table As Integer
Summary_table = 2

lastrow = Range("A" & Rows.Count).End(xlUp).Row

'loop trough all tickers
For i = 2 To lastrow
'check if the ticker name is unique, if not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Set the unique Ticker
    Ticker = ws.Cells(i + 1, 1).Value
    'Set Year open and Year close
    Year_close = ws.Cells(i, 6).Value
    Year_open = ws.Cells(ct, 3).Value
   
    'Set the Yearly Change
    Yearly_change = Year_close - Year_open
 
        'Set the Percent Change
        If Year_open = 0 Then
        Percent_change = 0
        Else
        Percent_change = Yearly_change / Year_open
        End If
        
   
    
      'Add to the Total Volume
    Total_vol = Total_vol + ws.Cells(i, 7).Value
    
    'Allocate the Ticker in the Summary Table
     ws.Cells(Summary_table, 9).Value = ws.Cells(i, 1).Value
    'Allocate the Total Volume in the Summary Table
     ws.Cells(Summary_table, 12).Value = Total_vol
    'Allocate the Percent Change in the Summary Table
     ws.Cells(Summary_table, 11).Value = Percent_change
    'Allocate the Yearly Change in the Summary Table
     ws.Cells(Summary_table, 10).Value = Yearly_change
    
    'Add one to the Summary Table
    Summary_table = Summary_table + 1
    
    If Yearly_change >= 0 Then
    ws.Cells(Summary_table, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(Summary_table, 10).Interior.ColorIndex = 3
    End If
        
    'Reset the Total Volume
    Total_vol = 0
    'Add one to obtain the next Yearly_open Value
    ct = i + 1
    Else
     'Add to the Total Volume
    Total_vol = Total_vol + ws.Cells(i, 7).Value

    End If
      
    If Yearly_change >= 0 Then
    ws.Cells(Summary_table, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(Summary_table, 10).Interior.ColorIndex = 3
    End If
        
    Next i
    
    Next ws
End Sub


