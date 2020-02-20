Sub HardSolution():
Dim currentws As Worksheet
For Each currentws In Worksheets 'loop through all worksheets
'Name variables and set their intial values
    Dim TickerName As String
    Dim YearOpen As Double
    YearOpen = 0
    Dim YearClose As Double
    YearClose = 0
    Dim YearlyChange As Double
    YearlyChange = 0
    Dim PercentChange As Double
    PercentChange = 0
    Dim TickerVolume As Double
    
    'Variables for hard part
    Dim MaxTickerName As String
    Dim MinTickerName As String
    Dim MaxPercent As Double
    Dim MinPercent As Double
    Dim MaxVolTicker As String
    Dim MaxVol As Double
    MaxVol = 0
    
    Dim SummaryTable As Long
    SummaryTable = 2
    
    Dim LastRow As Long
    Dim i As Long
    LastRow = currentws.Cells(Rows.Count, 1).End(xlUp).Row
' Set titles for summary table
    currentws.Range("I1").Value = "Ticker"
    currentws.Range("J1").Value = "Yearly Change"
    currentws.Range("K1").Value = "Percent Change"
    currentws.Range("L1").Value = "Total Stock Volume"
'Setting titles for the other table
    currentws.Range("O2").Value = "Greatest % Increase"
    currentws.Range("O3").Value = "Greatest % Decrease"
    currentws.Range("O4").Value = "Greatest Total Volume"
    currentws.Range("P1").Value = "Ticker"
    currentws.Range("Q1").Value = "Value"
    
    YearOpen = currentws.Cells(2, 3).Value ' initial value of open price
    
'Start of looping data
For i = 2 To LastRow
    'Ticker Symbol change check
    If currentws.Cells(i + 1, 1).Value <> currentws.Cells(i, 1).Value Then ' from nextcells activity
        TickerName = currentws.Cells(i, 1).Value ' placing data
        
        'calculating yearlychange
        YearClose = currentws.Cells(i, 6).Value ' getting data from this place
        YearlyChange = YearClose - YearOpen
          
        On Error Resume Next ' helps with my overflow error
    
        'calculating percent change
        PercentChange = (YearlyChange / YearOpen)
        currentws.Cells(SummaryTable, 11).Value = PercentChange
        
    
             'getting division by 0 error so to check that
             If YearOpen <> 0 Then
                 PercentChange = (YearlyChange / YearClose)
             Else
                  MsgBox ("fix manually")
             End If
    
         currentws.Range("K2").NumberFormat = "0.00%" 'change format to percentage
    
         'calculate total ticker volume
         TickerVolume = TickerVolume + currentws.Cells(i, 7).Value
    'placing values in summary table
    currentws.Cells(SummaryTable, 9).Value = TickerName
    currentws.Cells(SummaryTable, 10).Value = YearlyChange
    currentws.Cells(SummaryTable, 12).Value = TickerVolume
 
    
    'Conditional formatting for positive and negative values
    Dim rng As Range
    Dim LastRowYC As Long
    LastRowYC = currentws.Range("J" & Rows.Count).End(xlUp).Row
    
    Set rng = Range("J2" & LastRowYC)
        For Each cell In rng
            If (YearlyChange > 0) Then
                currentws.Range("J" & SummaryTable).Interior.ColorIndex = 4
            Else
                currentws.Range("J" & SummaryTable).Interior.ColorIndex = 3
            End If
        Next cell
      
    SummaryTable = SummaryTable + 1 'add to move onto next row of summary table
    YearOpen = currentws.Cells(i + 1, 3) ' to get next open price ticker value
    
    'setting variables to get last cell values of the percent change column
    Dim LastRowPercent As Long
    LastRowPercent = currentws.Cells(Rows.Count, 9).End(xlUp).Row
    Dim LastRowVol As Long
    LastRowVol = currentws.Cells(Rows.Count, 10).End(x1Up).Row
    
    'getting max and min values - info taken from google
     MaxPercent = Application.WorksheetFunction.Max(currentws.Range("K" & LastRowPercent))
     MinPercent = Application.WorksheetFunction.Min(currentws.Range("K" & LastRowPercent))
     MaxVol = Application.WorksheetFunction.Max(currentws.Range("L" & LastRowVol))
    
    ' Checking to get the highest and lowest value and placing them
    If (PercentChange > MaxPercent) Then
         MaxPercent = PercentChange
         MaxTickerName = TickerName
     ElseIf (PercentChange < MinPercent) Then
         MinPercent = PercentChange
         MinTickerName = TickerName
     End If
     
                       
     If (TickerVolume > MaxVol) Then
         MaxVol = TickerVolume
         MaxVolTicker = TickerName
     End If         
        
End If
    'getting ticker volume
    
    If currentws.Cells(i + 1, 1).Value <> currentws.Cells(i, 1).Value Then
       TickerVolume = 0
      Else
          TickerVolume = TickerVolume + currentws.Cells(i, 7).Value
      End If
      
      
      
    Next i
        'placing values in new table
        currentws.Range("Q2").Value = (CStr(MaxPercent) & "%")
        currentws.Range("Q3").Value = (CStr(MinPercent) & "%")
        currentws.Range("P2").Value = MaxTickerName
        currentws.Range("P3").Value = MinTickerName
        currentws.Range("Q4").Value = MaxVol
        currentws.Range("P4").Value = MaxVolTicker
        
    Next currentws
End Sub

