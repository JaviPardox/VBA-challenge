Attribute VB_Name = "Module2"
Sub StockMarketData()

    '****** THIS VBA SCRIPT ANALYZES REAL STOCK MARKET DATA ******'
    '
    '****** BY https://github.com/JaviPardox ******'

    Dim Ticker As String
    Dim TotalVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim stockCounter As Long
    Dim firstValue As Double
    Dim lastValue As Double
    Dim beginningYear As Long
    Dim endYear As Long
    
    
    'Extra variables for max and min purposes
    
    Dim maxPercentage As Double
    Dim percentage As Double
    Dim minPercentage As Double
    Dim maxTicker As String
    Dim minTicker As String
    Dim maxTotalVolume As Double
    Dim volTicker As String
    
    'WorkSheetCounter holds the value of the amount of worksheets in the file
    Dim workSheetCounter As Integer
    workSheetCounter = Application.Worksheets.Count
    
    'Initialize dates
    beginningYear = 20160101
    endYear = 20161230
    Ticker = "initializing"
    
    'For loop to iterate through all worksheets
    For j = 1 To workSheetCounter
        
        'Activate current worksheet
        Worksheets(j).Activate
        
        'Count the total amount of rows & initialize variables
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        stockCounter = 2
        maxPercentage = 0
        minPercentage = 0
        maxTotalVolume = 0
        
        'For loop to iterate through the worksheet
        For i = 2 To lastRow
        
            'Get the total volume by adding at each iteration the current value of volume
            TotalVolume = TotalVolume + ActiveSheet.Cells(i, 7).Value
            
            'Stop at the beginning of the year to store the ticker symbol
            'Store the open value as well
            If (ActiveSheet.Cells(i, 1).Value) <> Ticker Then
            
                Ticker = ActiveSheet.Cells(i, 1).Value          'get ticker symbol
                firstValue = ActiveSheet.Cells(i, 3).Value      'get the open value of the stock
                    
            End If
            
            'Stop at the end of the year to store the close value, calculate percentage,
            'Calculate yearly change, and total stock volume while keeping track of max and min values
            If (i > 2) And (i < (lastRow + 1)) And (ActiveSheet.Cells((i + 1), 1).Value <> Ticker) Then
            
                lastValue = ActiveSheet.Cells(i, 6).Value                           'get close value
                ActiveSheet.Cells(stockCounter, 9).Value = Ticker                   'write down ticker symbol
                ActiveSheet.Cells(stockCounter, 10).Value = lastValue - firstValue  'calculate difference
                
                'Checking if the first open value is 0
                If firstValue <> 0 Then
                
                    percentage = (((lastValue - firstValue) / firstValue) * 100)    'calculate percentage
                    ActiveSheet.Cells(stockCounter, 11).Value = percentage & "%"    '% symbol automatically rounds the number
                    
                    'keeping track of the maximum and minimum values
                    If percentage > maxPercentage Then
                    
                        maxPercentage = percentage
                        maxTicker = Ticker
                        
                    ElseIf percentage < minPercentage Then
                    
                        minPercentage = percentage
                        minTicker = Ticker
                        
                    End If
                    
                Else
                
                    ActiveSheet.Cells(stockCounter, 11).Value = 0
                    
                End If
                
                ActiveSheet.Cells(stockCounter, 12).Value = TotalVolume
                
                If TotalVolume > maxTotalVolume Then
                
                    maxTotalVolume = TotalVolume
                    volTicker = Ticker
                
                End If
                    
                'Color the cells
                If (ActiveSheet.Cells(stockCounter, 10).Value < 0) Then
                   
                    'Color red
                    ActiveSheet.Cells(stockCounter, 10).Interior.ColorIndex = 3
                   
                ElseIf (ActiveSheet.Cells(stockCounter, 10).Value > 0) Then
               
                    'Color green
                    ActiveSheet.Cells(stockCounter, 10).Interior.ColorIndex = 4
            
                End If
                
                TotalVolume = 0                   'reset the total volume before counting for the next symbol
                stockCounter = stockCounter + 1   'add 1 to the stack of unique symbols to keep track
            
            End If
            
            'Next row
            Next i
            
            'Finish worksheet iteration by writting the titles
            ActiveSheet.Cells(1, 9).Value = "Ticker"
            ActiveSheet.Cells(1, 10).Value = "Yearly Change"
            ActiveSheet.Cells(1, 11).Value = "Percent Change"
            ActiveSheet.Cells(1, 12).Value = "Total Stock Volume"
            ActiveSheet.Cells(1, 16).Value = "Ticker"
            ActiveSheet.Cells(1, 17).Value = "Value"
            ActiveSheet.Cells(2, 15).Value = "Greatest % Increase"
            ActiveSheet.Cells(3, 15).Value = "Greatest % Decrease"
            ActiveSheet.Cells(4, 15).Value = "Greatest Total Volume"
            ActiveSheet.Cells(2, 16).Value = maxTicker
            ActiveSheet.Cells(3, 16).Value = minTicker
            ActiveSheet.Cells(4, 16).Value = volTicker
            ActiveSheet.Cells(2, 17).Value = maxPercentage & "%"
            ActiveSheet.Cells(3, 17).Value = minPercentage & "%"
            ActiveSheet.Cells(4, 17).Value = maxTotalVolume
            'AutoFit to format cells
            ActiveSheet.Range("J1:L1").Columns.AutoFit
            ActiveSheet.Range("O2:O4").Columns.AutoFit
            
        'Next worksheet
        Next j
    
End Sub


