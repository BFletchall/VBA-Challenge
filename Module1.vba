Attribute VB_Name = "Module1"
Sub stock_challenge():

'populate on all worksheets
    For Each ws In Worksheets
    
    Dim lastrow As Long
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
'column and row headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Declare variables for ticker, yearly change, percent change, total stock  volume
Dim Ticker As String
Dim percentchng As Double
Dim i As Long
Dim j As Long
Dim Total As Double
Dim Start As Double
Dim Change As Double
Dim endprice As Double
Dim startprice As Double

j = 0
Total = 0
Start = 2
Change = 0
startprice = ws.Range("C2").Value

'loop for populating ticker symbol column summarize from col A

    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Total = Total + ws.Cells(i, 7).Value
                   
           Ticker = ws.Cells(i, 1).Value
           Change = ws.Cells(i, 6).Value - startprice
                      
            If startprice = 0 Then
                percentchng = 0
            Else
                percentchng = (Change / startprice)
            End If
           
           startprice = ws.Range("C" & i + 1)
                                    
           If Total = 0 Then
           Change = 0
           percentchng = "%" & 0
           
           Else
            If ws.Cells(Start, 3) = 0 Then
              For find_value = Start To i
                If ws.Cells(find_value, 3).Value <> 0 Then
                   Start = find_value
                   Exit For
                End If
               Next find_value
             End If
                
    ' next ticker
    Start = i + 1
    
    ws.Range("I" & 2 + j).Value = Ticker
    ws.Range("J" & 2 + j).Value = Change
    ws.Range("J" & 2 + j).NumberFormat = "0.00"
    ws.Range("K" & 2 + j).Value = percentchng
    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
    ws.Range("L" & 2 + j).Value = Total
    
 'conditional color red for neg and green for pos
           
        If Change > 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Change < 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
        End If
        
           End If
           
 ' reset variables for new stock ticker
           Total = 0
           Change = 0
           j = j + 1
                      
        'if ticker is the same
        Else
            Total = Total + ws.Cells(i, 7)
          
        End If
        
     Next i
     
' take the max and min and place them in a separate part in the worksheet
    
    Dim increase_number As Long
    Dim decrease_number As Long
    Dim volume_number As Long
    Dim maxticker As String
    Dim minticker As String
    Dim maxvol As String
    
' Calculate
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
    
'Find row numbers
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
    
    maxticker = Cells(increase_number + 1, 9).Value
    minticker = Cells(decrease_number + 1, 9).Value
    maxvol = Cells(volume_number + 1, 9).Value
    
    ws.Range("P2").Value = maxticker
    ws.Range("P3").Value = minticker
    ws.Range("P4").Value = maxvol
    
    
    Next ws

End Sub

