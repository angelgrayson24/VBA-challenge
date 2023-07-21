Sub stocks():
    ' Introducing our variables to excel
    Dim ticker As String
    Dim summaryLocation As Integer
    Dim LastRow As Long
    Dim Gti As Double
    Dim Gtd As Double
    Dim Gtv As Double
    Dim Gtit As String
    Dim Gtdt As String
    Dim Gtvt As String
    Dim ws As Worksheet
    Dim yearlychange As Double
    Dim openingprice As Double
    Dim closingprice As Double
    Dim Yearstartrow As Long
    Dim YearEndrow As Long
    Dim percentagechange As Double
    Dim Totalvolume As Double
    Dim volume As Double
    
     For Each ws In Worksheets
    summaryLocation = 2
    Gti = 0
    Gtd = 0
    Gtv = 0
    Gtit = ""
    Gtdt = ""
    Gtvt = ""
    
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
     openingprice = ws.Cells(2, 3).Value
     ws.Cells(1, 9).Value = "Ticker"
    
     ws.Cells(1, 10).Value = "Yearly Change"
     
     ws.Cells(1, 11).Value = "Percentage Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
    
  
    For i = 2 To LastRow:
        
        ' This if statement checks to see if this is the first row with this ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
        
        
           
            ' If this is the last row, we add the ticker to the cell
             ws.Cells(summaryLocation, 9).Value = ticker
            closingprice = ws.Cells(i, 6).Value
            yearlychange = (closingprice - openingprice)
             ws.Cells(summaryLocation, 10).Value = yearlychange
         ' Format yearly change
            If yearlychange < 0 Then
                ws.Cells(summaryLocation, 10).Interior.ColorIndex = 3
            ElseIf yearlychange > 0 Then
                ws.Cells(summaryLocation, 10).Interior.ColorIndex = 4
          End If
        
            
             percentagechange = (yearlychange / openingprice)
             
             openingprice = ws.Cells(i + 1, 3).Value
             ws.Cells(summaryLocation, 11).Value = FormatPercent(percentagechange, 2)
             
               If percentagechange > Gti Then
            Gti = percentagechange
            Gtit = ws.Cells(i, 1).Value
     
     
     ElseIf percentagechange < Gtd Then
            Gtd = percentagechange
            Gtdt = ws.Cells(i, 1).Value
    ElseIf Totalvolume > Gtv Then
            Gtv = Totalvolume
            Gtvt = ws.Cells(i, 1).Value
    End If
           
            ws.Cells(summaryLocation, 12).Value = Totalvolume
            summaryLocation = summaryLocation + 1
            
 End If
     
                volume = ws.Cells(i, 7).Value
                Totalvolume = Totalvolume + volume

    
        Next i
   
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
    
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("P2").Value = Gtit
     ws.Range("Q2").Value = Gti
     
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("P3").Value = Gtdt
     ws.Range("Q3").Value = Gtd
     
     ws.Range("O4").Value = "Greatest Total Volume"
     ws.Range("P4").Value = Gtvt
     ws.Range("Q4").Value = Gtv

    Next ws
    
            
End Sub