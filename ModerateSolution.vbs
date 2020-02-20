MODERATE SOLUTION - FINALLLLLLLLLL

Sub ModerateSolution():
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
    
    YearOpen = currentws.Cells(2, 3).Value ' initial value of open price
    
'Start of looping data
For i = 2 To LastRow
    'Ticker Symbol change check
    If currentws.Cells(i + 1, 1).Value <> currentws.Cells(i, 1).Value Then ' from nextcells activity
        TickerName = currentws.Cells(i, 1).Value ' placing data
        
        'calculating yearlychange and percent change
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
    'reset values
    YearlyChange = 0
    PercentChange = 0
    YearClose = 0
    YearOpen = currentws.Cells(i + 1, 3) ' to get next open price ticker value
    
End If
    If currentws.Cells(i + 1, 1).Value <> currentws.Cells(i, 1).Value Then
       TickerVolume = 0
      Else
          TickerVolume = TickerVolume + currentws.Cells(i, 7).Value
      End If
      
    Next i
        Next currentws
End Sub


