Attribute VB_Name = "Module1"
Sub StockSummary():

    'declare ws as a worksheet variable
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        'populate headers for summary
        ws.Range("J1").Value = "ticker"
        ws.Range("K1").Value = "yearly change"
        ws.Range("L1").Value = "percent change"
        ws.Range("M1").Value = "total volume"
    
        'set initial var to store ticker symbol
        Dim ticker As String
    
        'set variable to store opening and closing price
        Dim price_open As Double
        Dim price_close As Double
        price_open = ws.Cells(2, 3).Value
        
        'set variable to store calculated value for yearly change
        Dim yearly_change As Double
        
        'set variable to hold total volume for each ticker
        Dim volume As Variant
        volume = 0
    
        'keeps track of next available row in summary
        Dim summary_row As Integer
        summary_row = 2
        
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'detect ticker symbol change in ws then populates summary
        For i = 2 To lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'set value for ticker
                ticker = ws.Cells(i, 1).Value
                'populate ticker in summary
                ws.Range("J" & summary_row).Value = ticker
                
                'store value for closing price
                price_close = ws.Cells(i, 6).Value
                'Calculate yearly change
                yearly_change = price_close - price_open
                'populate yearly change in summary
                ws.Range("K" & summary_row).Value = yearly_change
                
                    '--------------------------------------------------------------------
                    'format yearly change column in summary so negative change is red and positive change is green
                    If yearly_change < 0 Then
                        ws.Range("K" & summary_row).Interior.ColorIndex = 3
                    
                    Else
                        ws.Range("K" & summary_row).Interior.ColorIndex = 4
                    
                    End If
                    '--------------------------------------------------------------------
                
                'populate percent change
                ws.Range("L" & summary_row).Value = yearly_change / price_open

                'reset price_open to the next ticker open price
                price_open = ws.Cells(i + 1, 3).Value
                
                'adds up total volume
                volume = volume + ws.Cells(i, 7).Value
                'populate total volume in summary
                ws.Range("M" & summary_row).Value = volume
                
                'keeps track of next available row in summary
                summary_row = summary_row + 1
                
                'reset total volume
                volume = 0
                
            Else
                volume = volume + ws.Cells(i, 7).Value
                
            End If
            
            
        Next i
        
        '---------------------------------------------------------------------
        'bonus summary
        'prints labels for bonus summary
        ws.Range("P2").Value = "Greatest % increase"
        ws.Range("P3").Value = "Greatest % decrease"
        ws.Range("P4").Value = "Greatest total value"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        
        'declares variables to find and hold approporiate values for second summary chart
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim max_volume As Variant
        
        'finds summary values within respective columns
        max_increase = Application.WorksheetFunction.Max(ws.Range("L:L"))
        max_decrease = Application.WorksheetFunction.Min(ws.Range("L:L"))
        max_volume = Application.WorksheetFunction.Max(ws.Range("M:M"))
        
        'Loops through first summary table
        For r = 2 To lastrow
            'fetches and populates max increase and max decrease ticker and value
            If ws.Cells(r, 12).Value = max_increase Then
                ws.Range("Q2").Value = ws.Cells(r, 10).Value
                ws.Range("R2").Value = max_increase
                
            ElseIf ws.Cells(r, 12).Value = max_decrease Then
                ws.Range("Q3").Value = ws.Cells(r, 10).Value
                ws.Range("R3").Value = max_decrease
                
            End If
            
            'fetches and populates max volume ticker and value
            If ws.Cells(r, 13).Value = max_volume Then
                ws.Range("Q4").Value = ws.Cells(r, 10).Value
                ws.Range("R4").Value = max_volume
            
            End If
        
        Next r
        
        'formats percent cells
        ws.Range("L:L").NumberFormat = "0.00%"
        ws.Range("R2:R3").NumberFormat = "0.00%"
        
        'autofit columns
        ws.Columns("A:R").AutoFit
        
      
    Next ws
    

End Sub


   
  
