VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'assign individual stock
Dim ticker As String
'assign annual change between opening and closing price
Dim changeinprice As Double
Dim openingprice As Double
Dim closingprice As Double
'assign annual percent change in price between opening and closing price
Dim percentchangeinprice As Double
'assign total stock volume for an individual stock
Dim totalstockvolume As Double

Sub stockmarket()
'create a script to loop through each worksheet with 1 click

For Each ws In Worksheets

'activate sheet
Set ws = ActiveSheet



'add summary table headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"
ws.Range("K1").Value = "Yearly $ change"
ws.Range("L1").Value = "Yearly % change"

'Create a script that will loop through all stocks for one year:

  'Keep track of the location for each stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  

  ' Loop through all stock data points:
  
  'Set initial variable for holding the total stock volume to 0
   totalstockvolume = 0
   openingprice = ws.Cells(2, 3).Value
  
  'auto-determine last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    For i = 2 To LastRow
    
    
      ' Check if we are still within the same stock ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the stock reference
      ticker = ws.Cells(i, 1).Value
      
      'set the closing stock price value
      closingprice = ws.Cells(i, 6).Value

      ' Add to the stock volume total
      totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value

      ' Print the stock ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker

      ' Print the total stock volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = totalstockvolume
      
      'calculate and print the annual change in price to the summary table
      changeinprice = closingprice - openingprice
      ws.Range("K" & Summary_Table_Row).Value = changeinprice
    
        
      'calculate and print the percentage change in annual price to the summary table
      percentagechangeinprice = changeinprice / openingprice
      On Error Resume Next
      ws.Range("L" & Summary_Table_Row).Value = FormatPercent(percentagechangeinprice, 2)
        
        
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total stock variables for next stock
      totalstockvolume = 0
      'Set the opening price for the next stock
      openingprice = ws.Cells(i + 1, 3).Value

    ' If the cell immediately following a row is the same stock ticker...
    Else

      ' Add to the total stock amount
      totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value

    End If

  Next i
  
  'working with summary table
  'format the annual change columns so positive change is green and negative change is red
    Dim SummaryTableLastRow
    SummaryTableLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As Double
    
    For i = 2 To SummaryTableLastRow
        
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.Color = RGB(0, 235, 0)
            Else: ws.Cells(i, 11).Interior.Color = RGB(235, 0, 0)
            End If
            If ws.Cells(i, 12).Value > 0 Then
                ws.Cells(i, 12).Interior.Color = RGB(0, 235, 0)
            Else: ws.Cells(i, 12).Interior.Color = RGB(235, 0, 0)
            End If
            
    Next i
    
    'create additional summary table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume of Trades"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'calculate greatest percent increase in summary
    greatestincrease = WorksheetFunction.Max(ws.Range("L2:L" & SummaryTableLastRow))
    ws.Range("Q2").Value = FormatPercent(greatestincrease, 2)
    'calculate greatest percent decrease in summary
    greatestdecrease = WorksheetFunction.Min(ws.Range("L2:L" & SummaryTableLastRow))
    ws.Range("Q3").Value = FormatPercent(greatestdecrease, 2)
    'calculate greatest total volume in summary
    greatestvolume = WorksheetFunction.Max(ws.Range("J2:J" & SummaryTableLastRow))
    ws.Range("Q4").Value = greatestvolume
    
    'vlookup the stock tickers for each of the new summary values and print to column P
    
    
 'activ
ActiveSheet.Next.Activate

Next ws

End Sub

