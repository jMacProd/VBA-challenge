Attribute VB_Name = "Module5"
Sub WorksheetLoop()

                 'tip for me - marco relies on excel sheet already sorted by ticker then by date
        
        'set up worksheet loop
        For s = 1 To ActiveWorkbook.Worksheets.Count
        Sheets(s).Activate
        'Enter code to be run on each worksheet here
   

'create summary table header
'-----------------------------
Range("J1").Value = "Ticker"
Range("k1").Value = "Yearly Change"
Range("L1").Value = "Percentage Change"
Range("M1").Value = "Total Stock Volume"

'set row for summary value
Dim summaryrow As Integer
summaryrow = 2

'create greatest table
'-------------------------------
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'add tickers, volumes, pricechange and percentage to summarytable
'-----------------------------------
'set totalvolume as variable
Dim totalvolume As Double
totalvolume = 0

'set ticker variable
Dim ticker As String

'Sent pricechange amounts as variable
Dim openprice As Variant
Dim closeprice As Variant
Dim change As Variant
Dim percentchange As Variant

'Set last row count
Dim LastRow As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row

    'start looping to summary table values
    For I = 2 To LastRow

        'ticker value
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        ticker = Cells(I, 1).Value
    
        'track totalvolume
        totalvolume = totalvolume + Cells(I, 7).Value
    
        'add ticker to summary table
        Range("j" & summaryrow).Value = ticker
    
        'add totalvolume to summary table
        Range("M" & summaryrow).Value = totalvolume
    
        summaryrow = summaryrow + 1
    
        'reset tickervolume
        totalvolume = 0
  
        Else
    
        totalvolume = totalvolume + Cells(I, 7).Value
    
        End If
    
    Next I

'calculating change from open value to close valur
'------------------------------------

'Need to reset summaryrow
summaryrow = 2

   'loop through calculating change
    For j = 2 To LastRow
        
    
        'set open value
        If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
        openprice = Cells(j, 3).Value
         
        'set closed value
        ElseIf Cells(j + 1, 1).Value <> Cells(j, 1) Then
        closeprice = Cells(j, 6).Value
 
        'set difference
        change = closeprice - openprice
        Cells(summaryrow, 11).Value = change
    
        'overcome divided by 0 error
            'Advice from askclass -  ' encounter divide by zero error in the % change formula,
                'the way to work this out is to equate the % change to 0 if the divisor is 0
        If change = 0 Or openprice = 0 Then
        percentchange = "Nil"
        Else
        
        percentchange = change / openprice
        Cells(summaryrow, 12).Value = FormatPercent(percentchange)
        End If
    
        'Setting conditional colours for change
        If Cells(summaryrow, 11).Value < 0 Then
        Cells(summaryrow, 11).Interior.ColorIndex = 3
        
        ElseIf Cells(summaryrow, 11).Value > 0 Then
        Cells(summaryrow, 11).Interior.ColorIndex = 4
        
        Else
        Cells(summaryrow, 11).Interior.ColorIndex = xlNone
        End If
    
        'to get to next line in summary table
        summaryrow = summaryrow + 1
      
        End If
   
  
   Next j


'setting the greatest table
'--------------------------
Dim rng As Range
Dim MaxVol As Double
Dim MaxIncrease As Double
Dim MaxDecrease As Double

'so formula only looks in the correct number of rows for each sheet
LastRow = Sheets(s).Range("M" & Rows.Count).End(xlUp).Row

    'set maxvolume
    'Set range from which to determine smallest value
    Set rng = Sheets(s).Range("M2:M3" & LastRow)

    'Worksheet function MAX returns the largest value in a range
    MaxVol = Application.WorksheetFunction.Max(rng)
    Range("Q4").Value = MaxVol

    'Set max percentage increase and descrease
    LastRow = Sheets(s).Range("L" & Rows.Count).End(xlUp).Row
    Set rng = Sheets(s).Range("L2:L3" & LastRow)

    MaxIncrease = Application.WorksheetFunction.Max(rng)
    Range("Q2").Value = FormatPercent(MaxIncrease)

    MaxDecrease = Application.WorksheetFunction.Min(rng)
    Range("Q3").Value = FormatPercent(MaxDecrease)

'To go to next worksheet
Next s
        
End Sub
