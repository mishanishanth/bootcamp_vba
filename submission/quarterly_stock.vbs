Sub stockdata()
'declare variables

    Dim ticker As String
    Dim next_ticker As String
    Dim volume As Double
    Dim total_volume As Double
    Dim rowcount As Long
    Dim leaderboardrow As Integer
    Dim lastoccurence As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim lastRowNum As Double
    Dim quarterlychange As Double
    Dim firstrow As Double
    Dim percent_change As Double
    Dim rounded_percent As Double
    Dim leaderboardrowcount As Integer
    Dim pcvalue As Double
    Dim pcvaluenext As Double
    Dim newleaderboardrow As Integer
    Dim tmppcvaluenext As Double
    Dim ticker_next As String
    Dim tmpvolume As Double
    Dim highestvolume As Double
    Dim ws As Worksheet
    
     For Each ws In ThisWorkbook.Worksheets
            ws.Activate
    'Setting the title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Value"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % increase"
    Range("N3").Value = "Greatest % decrease"
    Range("N4").Value = "Greatest Total Volume"
    
    
    
    total_volume = 0
    leaderboardrow = 2
    firstrow = 2
           
    lastRowNum = Cells(Rows.Count, 1).End(xlUp).Row
               
    'loop through the rows
        
    For rowcount = 2 To lastRowNum
    
    'extract the values from worksheet to variables
        open_price = Cells(firstrow, 3)
        close_price = Cells(rowcount, 6)
        ticker = Cells(rowcount, 1).Value
        next_ticker = Cells(rowcount + 1, 1).Value
        volume = Cells(rowcount, 7)
    
    'check for condition
    
    If (ticker <> next_ticker) Then
    total_volume = total_volume + volume
    
    'calculate the quarterly change
    quarterlychange = Cells(rowcount, 6) - open_price
        
    'calculate percent change
    percent_change = quarterlychange / open_price
   ' rounded_percent = Application.WorksheetFunction.RoundUp(percent_change, 2)
    
    firstrow = rowcount + 1 'to find the firstrow of every new ticker
    
    'write to leaderboard
    
    Cells(leaderboardrow, 9).Value = ticker
    Cells(leaderboardrow, 10).Value = quarterlychange
    Cells(leaderboardrow, 11).Value = FormatPercent(percent_change)
    Cells(leaderboardrow, 12).Value = total_volume
    
    For leaderboardrowcount = 2 To 1501
       
        If Cells(leaderboardrowcount, 10).Value > 0 Then
        
            Cells(leaderboardrowcount, 10).Interior.ColorIndex = 4
       ElseIf Cells(leaderboardrowcount, 10).Value < 0 Then
            Cells(leaderboardrowcount, 10).Interior.ColorIndex = 3
       Else
            Cells(leaderboardrowcount, 10).Interior.ColorIndex = 2
       
       End If
       
       Next leaderboardrowcount
    
    
    total_volume = 0
    leaderboardrow = leaderboardrow + 1
    
    
    Else
    
    total_volume = total_volume + volume
    
    End If
    
    Next rowcount
    
    newleaderboardrow = 2
     tmppcvaluenext = 0
     ticker_next = ""
   
    For leaderboardrowcount = 2 To leaderboardrow
    
      ticker = Cells(leaderboardrowcount, 9).Value
       ' MsgBox (ticker)
     pcvalue = Cells(leaderboardrowcount, 11).Value
     
     If pcvalue > tmppcvaluenext Then
            tmppcvaluenext = pcvalue
            ticker_next = ticker
           ' MsgBox (ticker_next)
            
     End If
       
      Cells(newleaderboardrow, 15) = ticker_next
      Cells(newleaderboardrow, 16) = FormatPercent(tmppcvaluenext)
                 
      Next leaderboardrowcount
      
     tmppcvaluenext = pcvalue
      
      For leaderboardrowcount = 2 To leaderboardrow
    
      ticker = Cells(leaderboardrowcount, 9).Value
       
     pcvalue = Cells(leaderboardrowcount, 11).Value
     
     If pcvalue < tmppcvaluenext And pcvalue <= tmppcvaluenext Then
           tmppcvaluenext = pcvalue
            ticker_next = ticker
          
                       
     End If
            
           Next leaderboardrowcount
             
                      
               newleaderboardrow = newleaderboardrow + 1
               
              Cells(newleaderboardrow, 15) = ticker_next
               Cells(newleaderboardrow, 16) = FormatPercent(tmppcvaluenext)
               
           tmpvolume = 0
           
               
       For leaderboardrowcount = 2 To leaderboardrow
    
      ticker = Cells(leaderboardrowcount, 9).Value
       ' MsgBox (ticker)
     highestvolume = Cells(leaderboardrowcount, 12).Value
     
     If highestvolume > tmpvolume Then
            tmpvolume = highestvolume
            ticker_next = ticker
           ' MsgBox (ticker_next)
            
     End If
       
      
                 
      Next leaderboardrowcount
               
               
      newleaderboardrow = newleaderboardrow + 1
               
              Cells(newleaderboardrow, 15) = ticker_next
               Cells(newleaderboardrow, 16) = tmpvolume
               
Next ws
   
        
End Sub

