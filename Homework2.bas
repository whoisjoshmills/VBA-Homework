Attribute VB_Name = "Module1"
Sub StockTracker()

'VARIABLES

'Variable to count number of worksheets in a workbook dim
Dim WS_Count As Long

'SUMMARY TABLE dims
Dim i As Long
Dim x As Long
Dim Ticker As String
Dim LastRow As String
Dim YearlyChangeStart As Double
Dim YearlyChangeFinish As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double

'BONUS dims
Dim LargestPercent As Double
Dim SmallPercent As Double
Dim LargestTotalVol As Double
Dim LargestPercentTicker As String
Dim SmallPercentTicker As String
Dim LargestTotalVolTicker As String
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Determine number of worksheets in current workbook
WS_Count = ActiveWorkbook.Worksheets.Count



'Iterate through each worksheet and complete desired task
For j = 2 To WS_Count
    Worksheets(j).Activate
      
    'SUMMARY TABLE - Creating the summary table headers
    ActiveSheet.Cells(1, 9).Value = "Ticker"
    ActiveSheet.Cells(1, 10).Value = "Yearly Change"
    ActiveSheet.Cells(1, 11).Value = "Percent Change"
    ActiveSheet.Cells(1, 12).Value = "Total Stock Volume"
        
    'BONUS TABLE - Creating the bonus table headers
    ActiveSheet.Cells(1, 16).Value = "Ticker"
    ActiveSheet.Cells(1, 17).Value = "Value"
    ActiveSheet.Cells(2, 15).Value = "Greatest % Increase"
    ActiveSheet.Cells(3, 15).Value = "Greatest % Decrease"
    ActiveSheet.Cells(4, 15).Value = "Greatest Total Volume"
          
    'i tracks what row of my Summary Table I'm trying to fill out
    i = 2
    
    'x will be used to know which rows we will need to sum to calculate TOTAL STOCK VOLUME. I start with 2 because I want to save the first ticker's january 1st starting bid (cell G2)
    x = 2
    
    'YEARLY CHANGE - Storing the first stock's open price on Jan 1 to be used later to find the difference with the year end price in the MAIN FOR LOOP
    YearlyChangeStart = Cells(i, 3).Value
    
    'FOR LOOP - Finding and storing the number of rows in the data set. Used in the MAIN FOR LOOP as a stop point to the For loop
     
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Range("A1").End(xlDown).Select
    'LastRow = ActiveCell.Row
    
    'BONUS - defining the bonus related variables
    LargestPercent = 0
    SmallPercent = 0
    LargestTotalVol = 0
'-----------------------------------------------------------------------------------------------------------------------------------------

    'MAIN FOR LOOP
    'The For loop will go through the ticker symbol and determine if the ticker has changed from row to row. If it has changed, the Ticker, Yearly Change, Percent Change, and Total Stock Volume will be calculated and placed into their respective cells for that line item
    For a = 2 To LastRow
        If ActiveSheet.Cells(a + 1, 1).Value <> Cells(a, 1).Value Then
            
            'TICKER - If ticker has changed, then place new ticker into summary table
            Ticker = ActiveSheet.Cells(a, 1).Value
            ActiveSheet.Cells(i, 9).Value = Ticker
               
               
               
            'YEARLY CHANGE - If ticker has changed, calculate yearly difference and place into summary table
            YearlyChangeFinish = Cells(a, 6).Value
            Cells(i, 10).Value = YearlyChangeFinish - YearlyChangeStart
            
            'Conditionally formatting cells
            If ActiveSheet.Cells(i, 10) > 0 Then
                ActiveSheet.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf ActiveSheet.Cells(i, 10) < 0 Then
                ActiveSheet.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
            End If
                    
                    
                    
            'PERCENT CHANGE - If ticker has changed, calculate percent difference and place into summary table and format to match percentage with 2 decimal places
            'If statment is used to prevent errors from dividing by 0
            If YearlyChangeStart = 0 Then
                PercentChange = 0
                ActiveSheet.Cells(i, 11).Value = PercentChange
                ActiveSheet.Cells(i, 11).NumberFormat = "0.00%"
            
            'Input of percent change into summary table and also changing number format
            Else
                PercentChange = (YearlyChangeFinish - YearlyChangeStart) / YearlyChangeStart
                ActiveSheet.Cells(i, 11).Value = PercentChange
                ActiveSheet.Cells(i, 11).NumberFormat = "0.00%"
            End If
            YearlyChangeStart = Cells(a + 1, 3).Value
            
            
            
            'TOTAL STOCK VOLUME - If ticker has changed, calculate the difference between the stock's <vol> on Dec 30th and the stock's <vol> on Jan 1st
            TotalStockVolume = Application.Sum(Range(Cells(x, 7), Cells(a, 7)))
            ActiveSheet.Cells(i, 12).Value = TotalStockVolume
            x = a + 1
            i = i + 1
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
            'BONUS - Check to see if new percentage is largest, smallest, or the volume was the greatest
            
            'Stock with the highest percent increase
            If PercentChange > LargestPercent Then
                LargestPercent = PercentChange
                LargestPercentTicker = Ticker
            End If
            
            'Stock with the highest percent increase
            If PercentChange < SmallPercent Then
                SmallPercent = PercentChange
                SmallPercentTicker = Ticker
            End If
            
            'Stock with the highest percent increase
            If TotalStockVolume > LargestTotalVol Then
                LargestTotalVol = TotalStockVolume
                LargestTotalVolumeTicker = Ticker
            End If
            
        End If
    Next a
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'BONUS TABLE - inputting in calculated values into bonus table
    
    'Input stock tickers for Bonus Table
    ActiveSheet.Cells(2, 16).Value = LargestPercentTicker
    ActiveSheet.Cells(3, 16).Value = SmallPercentTicker
    ActiveSheet.Cells(4, 16).Value = LargestTotalVolumeTicker
    
    'Input data values and format percentages
    ActiveSheet.Cells(2, 17).Value = LargestPercent
    ActiveSheet.Cells(2, 17).NumberFormat = "0.00%"
    ActiveSheet.Cells(3, 17).Value = SmallPercent
    ActiveSheet.Cells(3, 17).NumberFormat = "0.00%"
    ActiveSheet.Cells(4, 17).Value = LargestTotalVol
    
    'Format column lengths
    ActiveSheet.Cells.Select
    ActiveSheet.Cells.EntireColumn.AutoFit
    ActiveSheet.Cells(1, 1).Select

    'Clear stored values to start fresh on next worksheet
    LargestPercent = 0
    SmallPercent = 0
    LargestTotalVol = 0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Next worksheet
Next j

End Sub


' Questions to ask TA
' - How is the code going to be graded
