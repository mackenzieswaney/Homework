Sub Stock_Data()

' Loop through each worksheet
Dim ws As Worksheet
For Each ws In Worksheets

 ' Set an initial variable for holding the Ticker
  Dim Ticker As String
  Ticker = 0

  ' Set an initial variable for holding the Total Stock Volume
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  
  ' Create variable to loop through "last row"
  LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

  ' Loop through all Tickers
  For i = 2 To LastRow

    ' Check if we are still within the same Ticker, if it is not...
    If ws.Cells((i + 1), 1).Value <> ws.Cells(i, 1).Value Then

    ' Set ticker
    Ticker = ws.Cells(i, 1).Value

      ' Add to the Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Total_Stock_Volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total_Stock_Volume
      Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If

  Next i

      ' Reset the Total_Stock_Volume
      Total_Stock_Volume = 0
  
  Next ws

End Sub