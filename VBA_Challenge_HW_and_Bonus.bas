Attribute VB_Name = "Module1"
Sub Stocks()
    
   For Each ws In Worksheets
       
    'Set column headers
     ws.Range("J1").Value = "Ticker"
     ws.Range("K1").Value = " Total Volume"
     ws.Range("L1").Value = "Open Price"
     ws.Range("M1").Value = "Close Price"
     ws.Range("N1").Value = "Yearly Change"
     ws.Range("O1").Value = "Percentage Change"
     ws.Range("R1").Value = "Ticker"
     ws.Range("S1").Value = "Value"
     ws.Range("Q2").Value = "Greatest %increase"
     ws.Range("Q3").Value = "Greatest %decrease"
     ws.Range("Q4").Value = "Greatest total volume"
    
   
   
    'Set variable for holding ticker symbol
     Dim ticker As String
    
    'Set variable for counting volume & number of entries per ticker
     Dim volumecount As Double
     
     Dim number_entries As Integer
     
    'Set variables for opening & closing price
     Dim opening_price As Long
     Dim closing_price As Long
         
    'Syntx to count last row of stocks
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Initialize volume count to zero
     volumecount = 0
     
    'Initialize number of entries per ticker symbol
     number_entries = 0
    
    'Set variable to keep track of stock in table
     Dim Stock_summary_table As Integer
     Stock_summary_table = 2
     
    'Set variable to calculate stock value percent change
     Dim percent_change As Long
          
    'Set variable to calclate difference of opening price to closing price
     Dim open_minus_close As Long
       
       
        For i = 2 To LastRow
                    
            'Check if we are still within the same ticker, if it is not then...
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set the ticker symbol
                 ticker = ws.Cells(i, 1).Value
             
                'Set the closing price
                 closing_price = ws.Cells(i, 6).Value
             
                'Set the opening price
                 opening_price = ws.Cells(i - number_entries, 3).Value
                 
                 'Conditional statement to handle Zero values for opening price
                  If opening_price = 0 Or closing_price = 0 Then
                     percent_change = 0
                 
                  Else
                 
                    'Percent change formula
                     percent_change = (opening_price - closing_price) / opening_price
             
                    'Open minus close formula
                     open_minus_close = opening_price - closing_price
                 
                    'Print stock value percent change
                     ws.Range("O" & Stock_summary_table).Value = (opening_price - closing_price) / opening_price
                 
                    'Add to the stock volume
                     volumecount = volumecount + Cells(i, 7).Value
            
                    'Print the stock brand in the summary table
                     ws.Range("J" & Stock_summary_table).Value = ticker
             
                    'Print the stock volume total in the summary table
                     ws.Range("K" & Stock_summary_table).Value = volumecount
                
                    'Print closing price of ticker symbol
                     ws.Range("M" & Stock_summary_table).Value = closing_price
             
                    'Print opening price of ticker symbol
                     ws.Range("L" & Stock_summary_table).Value = opening_price
                    
                  End If
                  
                             
                'Print opening price versus closing price delta
                 ws.Range("N" & Stock_summary_table).Value = open_minus_close
                          
                            
                'Format stock percent change cells to percent type
                 ws.Range("O" & Stock_summary_table).NumberFormat = "0.00%"
                          
             
                       'Conditional loop to indicate positive or negative delta between opeining and closing price
                        If open_minus_close >= 0 Then
                           ws.Range("N" & Stock_summary_table).Interior.ColorIndex = 4
                
                        Else
                           ws.Range("N" & Stock_summary_table).Interior.ColorIndex = 3
                                                                     
                        End If
            
                'Add 1 to the stock summary table row
                 Stock_summary_table = Stock_summary_table + 1
             
                'Reset the volumecount
                 volumecount = 0
             
                'Reset the number_entries count
                 number_entries = 0
             
            'If the cells immmediately following the preceeding row is the same ticker value
            Else
    
                'Add to the stock volume
                 volumecount = volumecount + ws.Cells(i, 7).Value
                 number_entries = number_entries + 1
                    
           
            End If
        
        
        Next i
    
      
  
    'Set variable for range to find max percent increase, max decrease, max volume
     Dim rng_max_percent As Range
     Dim rng_min_percent As Range
     Dim rng_max_volume As Range
          
    'Set variables for min & max
     Dim max_percent As Double
     Dim min_percent As Double
     Dim max_volume As Double
     
        
    'Set variable to find last row
     Dim last_row As Integer

    'Set range from which to determine largest value
     Set rng_max_percent = ws.Range("O:O")
     Set rng_min_percent = ws.Range("O:O")
     Set rng_max_volume = ws.Range("K:K")

    'Worksheet function MAX_Percent increase returns the largest value in a range
     max_percent = Application.WorksheetFunction.Max(rng_max_percent)
     
    'Worksheet function MAX decrease returns the largest value in a range
     min_percent = Application.WorksheetFunction.Min(rng_min_percent)
     
    'Worksheet function MAX volume returns the largest value in a range
     max_volume = Application.WorksheetFunction.Max(rng_max_volume)

    
        'For loop o find max value of column and return ticker symbol
         For i = 2 To LastRow
            'Conditional loop to find cell value equal to max increase and then print ticker from the ith row
             If ws.Cells(i, 15).Value = max_percent Then
                ws.Cells(2, 18).Value = ws.Cells(i, 10).Value
                
                'Print max percentage
                ws.Range("S2").Value = max_percent
                ws.Range("S2").NumberFormat = "0.00%"
      
             End If
 
         Next i
         
        'For loop to find max volume of column and return ticker symbol
         For i = 2 To LastRow
            'Conditional loop to find cell value equal to max and then print ticker from the ith row
             If ws.Cells(i, 11).Value = max_volume Then
                ws.Cells(4, 18).Value = ws.Cells(i, 10).Value
                
                'Print max percentage
                ws.Range("S4").Value = max_volume
                      
             End If
 
         Next i
         
         'For loop to find max decreasee of column and return ticker symbol
          For i = 2 To LastRow
            'Conditional loop to find cell value equal to max decrease and then print ticker from the ith row
             If ws.Cells(i, 15).Value = min_percent Then
                ws.Cells(3, 18).Value = ws.Cells(i, 10).Value
                
                'Print max percentage
                ws.Range("S3").Value = min_percent
                ws.Range("S3").NumberFormat = "0.00%"
      
             End If
 
         Next i

    Next ws
    
    
End Sub


