Attribute VB_Name = "Module4"
Sub stockdata()

    For Each ws In Worksheets

       
        Dim wsname As String
        Dim yearchange As Double
        Dim percentchange As Double
        Dim stockvolume As Double
        
        
        
        

    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        wsname = ws.Name
        Dim ticker As String
        Dim value_open As Double
        Dim value_close As Double
        Dim outputcounter As Double
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        

        ticker = ws.Cells(2, 1).Value
        value_open = ws.Cells(2, 3).Value
        value_close = ws.Cells(2, 6).Value
        volume = ws.Cells(2, 7).Value
        outputcounter = 2
        
        
       
       
        For a = 3 To lastrow
        
        
        
            If ws.Cells(a, 1).Value <> ticker Then
                ws.Cells(outputcounter, 9).Value = ticker
                value_close = ws.Cells(a - 1, 6)
                ws.Cells(outputcounter, 10).Value = (value_close - value_open)
                
                
                If (value_close - value_open) >= 0 Then
                Cells(outputcounter, 10).Interior.ColorIndex = 4
                
                Else: Cells(outputcounter, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(outputcounter, 11).Value = ((value_close - value_open) / value_open) * 100
                ws.Cells(outputcounter, 12).Value = volume
                
                
                ticker = ws.Cells(a, 1).Value
                value_open = ws.Cells(a, 3).Value
                value_close = ws.Cells(a, 6).Value
                volume = ws.Cells(a, 7).Value
                
                
                outputcounter = outputcounter + 1
                
                   
           
                
            Else: volume = ws.Cells(a, 7).Value + volume
            
            
          End If
            
            
         Next a
       
       Next ws
       
       
        
        
    

End Sub
