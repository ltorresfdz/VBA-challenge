Attribute VB_Name = "Module1"
Sub ticker_info()

Dim Current As Worksheet

For Each Current In Worksheets
  
  Dim ticker_name As String     ' Set an initial variable for holding the ticker name
  Dim ticker_volume As Variant  ' Set an initial variable for holding the total of volume per ticker
  ticker_volume = 0
  Dim Summary_Table_Row As Integer  ' Keep track of the location for each ticker name in the summary table
  Summary_Table_Row = 2
  Dim first_value As Double    ' set the  year's first value of each ticker
  Dim last_value As Double      ' set the  year's last value of each ticker
  first_value = Cells(2, 3)
  last_value = 0
  Dim Yearly_change As Double   'to calculate the yearly change
  Yearly_change = 0
  Dim Percent_change As Double    'to calculate the percent change
  Percent_change = 0
    Current.Cells(1, 9).Value = "Ticker"     'Print the headers
    Current.Cells(1, 10).Value = "Yearly Change"
    Current.Cells(1, 11).Value = "Percent Change"
    Current.Cells(1, 12).Value = "Total Stock Volume"
    Current.Cells(2, 15).Value = "Greatest % Increase"
    Current.Cells(3, 15).Value = "Greatest % Decrease"
    Current.Cells(4, 15).Value = "Greatest Total Volume"
    Current.Cells(1, 16).Value = "Ticker"
    Current.Cells(1, 17).Value = "Value"
 
    
  For i = 2 To Current.Cells(Rows.Count, 1).End(xlUp).Row     ' Loop through all tickers daily transactions
 
    If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then     ' Check if we are still within the same ticker, if it is not...
      last_value = Current.Cells(i, 6)
      Yearly_change = last_value - first_value
      If first_value = 0 Then
        Percent_change = 1
      Else
        Percent_change = Yearly_change / first_value
      End If
      first_value = Current.Cells(i + 1, 3)                        'modificar el first value para el siguiente ticker.
      ticker_name = Current.Cells(i, 1).Value                      ' Set the ticker name
      ticker_volume = ticker_volume + Current.Cells(i, 7).Value    ' Add to the ticker volume
      Current.Range("I" & Summary_Table_Row).Value = ticker_name   ' Print the ticker name in the Summary Table
      Current.Range("J" & Summary_Table_Row).Value = Yearly_change 'Print the Yearly change in the Summary Table
      If Yearly_change < 0 Then                             'Color red if negative, green if positive
        Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      Else
        Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
      Current.Range("K" & Summary_Table_Row).Value = Format(Percent_change, "Percent") 'Print the Percent of change in the Summary Table
      Current.Range("L" & Summary_Table_Row).Value = ticker_volume ' Print the ticker volume to the Summary Table
      Summary_Table_Row = Summary_Table_Row + 1            ' Add one to the summary table row
      ticker_volume = 0                                    ' Reset the ticker volume
      

    
    Else                                                    ' If the cell immediately following a row is the ticker name...

      ticker_volume = ticker_volume + Current.Cells(i, 7).Value     ' Add to the ticker volume

    End If

  Next i

'to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

    Dim Maximo As Variant
    Dim Minimo As Variant
    Dim Volumen_Mayor As Variant
    Dim fin As Integer
    fin = Current.Cells(Rows.Count, 11).End(xlUp).Row

    Maximo = Application.WorksheetFunction.Max(Current.Range("K2:K" & fin))
    Current.Range("Q" & 2).Value = Format(Maximo, "Percent")
    Minimo = Application.WorksheetFunction.Min(Current.Range("K2:K" & fin))
    Current.Range("Q" & 3).Value = Format(Minimo, "Percent")
    Volumen_Mayor = Application.WorksheetFunction.Max(Current.Range("L2:L" & fin))
    Current.Range("Q" & 4).Value = Volumen_Mayor
    
    For i = 2 To fin
        If Current.Cells(i, 11).Value = Maximo Then                  ' Check if the ticker is the Maximo
          Current.Cells(2, 16).Value = Current.Cells(i, 9)                   ' Retrieve the values associated with the winner and enter them into the winner's box.
        End If
        If Current.Cells(i, 11).Value = Minimo Then
            Current.Cells(3, 16).Value = Current.Cells(i, 9)
        End If
        If Current.Cells(i, 12).Value = Volumen_Mayor Then
            Current.Cells(4, 16).Value = Current.Cells(i, 9)
        End If
    Next i
    
    MsgBox (Current.Name)

 Next Current

End Sub


