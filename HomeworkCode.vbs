Sub alphabetical_testing()

 
  Dim ticker As String
  Dim Yearly_Change, days, annual_open, annual_close, percentage_change, MaxValue As Double
  Total_Stock_Volume = 0
  'Cells(2, 13) = Cells(2, 3).Value
  annual_open = 0
  annual_close = 0
  percentage_change = 0
  MaxValue = 0
   
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  
  ' Loop through all data
  For i = 2 To lastrow
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ticker = Cells(i, 1).Value
      
     'Total Stock Volumes
     
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
      'Yearly Change
      days = days + 1
      annual_open = Cells(i - days + 1, 3).Value
      annual_close = Cells(i, 6).Value
      Yearly_Change =  annual_close -annual_open
      
      
      '%Change
      
        If annual_open = 0 Then
            Percantage_Change = 0
        
            Else
            percentage_change = annual_close / annual_open - 1
             'Range("K2").NumberFormat = "0.00%"
        End If
      
      Range("I" & Summary_Table_Row).Value = ticker
      Range("J" & Summary_Table_Row).Value = Yearly_Change
      
            If Yearly_Change > 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
      Range("k" & Summary_Table_Row).Value = percentage_change
      Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      
      
      'Range("M" & Summary_Table_Row).Value = annual_open
      'Range("N" & Summary_Table_Row).Value = annual_Close
      
      
      
      Summary_Table_Row = Summary_Table_Row + 1
      Summary_Table_Row_Open = Summary_Table_Row_Open + 1
        Total_Stock_Volume = 0
        days = 0
        
    Else
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
       days = days + 1
       
        
    End If


  Next i
Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("l1").Value = "Total Stock Volume"
Range("p1").Value = "Ticker"
Range("q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Deacrease"
Range("O4").Value = "Greatest Total Volume"


    MaxValue = Application.WorksheetFunction.Max(Range("K:K"))
    Cells(2, 17).Value = MaxValue

    MinValue = Application.WorksheetFunction.Min(Range("K:K"))
    Cells(3, 17).Value = MinValue

        Range("Q2:Q3").NumberFormat = "0.00%"
        
    MaxVolume = Application.WorksheetFunction.Max(Range("l:l"))
    Cells(4, 17).Value = MaxVolume
        
 For i = 2 To lastrow
 
      If Cells(i, 11) = Range("q2") Then
        Range("p2").Value = Cells(i, 9).Value
        End If
        
       If Cells(i, 11) = Range("q3") Then
        Range("p3").Value = Cells(i, 9).Value
        End If
    
    If Cells(i, 12) = Range("q4") Then
        Range("p4").Value = Cells(i, 9).Value
        End If

    Next i
      
      
        
End Sub
