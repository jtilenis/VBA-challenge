Attribute VB_Name = "Module1"
Sub Ticker_Summary()

  ' Set an initial variable for holding the Ticker Symbol
  Dim Ticker_Symbol, Symbol_Dec, Symbol_Inc, Symbol_Vol As String

  ' Set an initial variables for holding last_row, Open_Price, Close_Price, pct_change, and stock_volume
  Dim Open_Price, Close_Price, pct_change, Grt_Inc, Grt_Dec As Double
  Dim stock_volume, Grt_Volume As Variant
  Dim last_row As Long
  last_row = 0
  stock_volume = 0
  Open_Price = 0
  Close_Price = 0
  Price_Change = 0
  pct_change = 0
  Grt_Inc = 0
  Grt_Dec = 0
  Grt_Volume = 0
   
  'Add summary table to each worksheet
  For Each WS In Worksheets
  'MsgBox (ws.Name)
  ' Keep track of the location for each Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Find last row
  last_row = WS.Cells(Rows.Count, 1).End(xlUp).Row
  
  'Create Summary Headings
  WS.Range("I1").Value = "Ticker"
  WS.Range("J1").Value = "Yearly Change"
  WS.Range("K1").Value = "Percent Change"
  WS.Range("L1").Value = "Total Stock Volume"
  WS.Range("L1").ColumnWidth = 11
  
  ' Loop through all Ticker Symbols
  For i = 2 To last_row
 'MsgBox (Ticker_Symbol)
  'Check if we are still within the same Ticker Symbol range, if it is not...
    'At beginning of Ticker Symbol range get oen price
    If WS.Cells(i - 1, 1).Value <> WS.Cells(i, 1).Value Then
          Open_Price = WS.Cells(i, 3).Value

    ' Check if we are still within the same Ticker Symbol range, if it is not...
    ElseIf WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

      ' Set the Brand name
      Ticker_Symbol = WS.Cells(i, 1).Value
      'If Ticker_Symbol = "PLNT" Then
      'MsgBox (Ticker_Symbol)
      'End If
      'MsgBox (Ticker_Symbol)
      'Get Close Price
      Close_Price = WS.Cells(i, 6).Value
      ' Add to stock volume
      stock_volume = stock_volume + WS.Cells(i, 7).Value

      ' Print the Ticker Symbol in the Summary Table
      WS.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      'Print Open Price to summary table
      'Print yearly price change
       Price_Change = Close_Price - Open_Price
      WS.Range("J" & Summary_Table_Row).Value = Price_Change
      'Set Cell color
      If Price_Change < 0 Then
        WS.Range("J" & Summary_Table_Row).Interior.Color = vbRed
      Else
        WS.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
      End If
      'Print Percent Change to summary table
      If (Price_Change = 0) Or (Open_Price = 0) Then
      pct_change = 0
      Else
      pct_change = Price_Change / Open_Price
      End If
      
      WS.Range("K" & Summary_Table_Row).Value = Format(pct_change, "0.00%")
      'Print the stock volume to the Summary Table
      WS.Range("L" & Summary_Table_Row).Value = stock_volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the stock volume
      stock_volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the stock volume
      stock_volume = stock_volume + WS.Cells(i, 7).Value

    End If

  Next i
  
  Next WS
  MsgBox ("Stock Summary Done")
  
  'Loop thru the worksheets and find Greatest Increase, Greatest Decrease, Greatest Volume
  For Each WS In Worksheets
    'Find last row in summary section
    last_row = WS.Cells(Rows.Count, 9).End(xlUp).Row
    
     'Initialize variables
     Grt_Inc = 0
     Grt_Dec = 0
     Grt_Volume = 0
    
    'Loop thru summary section
     For i = 2 To last_row
     
     Ticker_Symbol = WS.Range("I" & i).Value
     pct_change = WS.Range("K" & i).Value
     'Find Greatest % increase
     If Grt_Inc < pct_change Then
     Symbol_Inc = Ticker_Symbol
     Grt_Inc = pct_change
     End If
     'Find Greatest % decrease
     If Grt_Dec > pct_change Then
     Symbol_Dec = Ticker_Symbol
     Grt_Dec = pct_change
     End If
     stock_volume = WS.Range("L" & i).Value
     'Find greatest symbol
     If Grt_Volume < stock_volume Then
     Symbol_Vol = Ticker_Symbol
     Grt_Volume = stock_volume
     End If
     
     Next i
     
     'Column Headers
    WS.Range("O1").Value = "Ticker"
    WS.Range("P1").Value = "Value"
    WS.Range("N1").ColumnWidth = 20
    WS.Range("P1").ColumnWidth = 11
  'Detail
    'Greatest % increas
    WS.Range("N2").Value = "Greatest % increase"
    WS.Range("O2").Value = Symbol_Inc
    WS.Range("P2").Value = Format(Grt_Inc, "0.00%")
    'Greatest % decrease
    WS.Range("N3").Value = "Greatest % decrease"
    WS.Range("O3").Value = Symbol_Dec
    WS.Range("P3").Value = Format(Grt_Dec, "0.00%")
   'Greatest total volume
    WS.Range("N4").Value = "Greatest total volume"
    WS.Range("O4").Value = Symbol_Vol
    WS.Range("P4").Value = Grt_Volume
  
     
  
  Next WS
      
  MsgBox ("Challenge done")

End Sub


