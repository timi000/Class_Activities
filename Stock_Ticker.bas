Attribute VB_Name = "Module2"
Sub Coderunner()
    
    For Each ws In Worksheets
    
        ws.Select
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Yearly Change"
        ws.Cells(1, 17).Value = "Percent Change"
        ws.Cells(1, 18).Value = "Total Stock Volume"
        ws.Cells(1, 22).Value = "Ticker"
        ws.Cells(1, 23).Value = "Volume"
        ws.Cells(2, 21).Value = "Greatest % Increase"
        ws.Cells(3, 21).Value = "Greatest % Decrease"
        ws.Cells(4, 21).Value = "Greatest Total Volume"
        
        
        
        
        
        
       
        Call Ticker_Volume
        
        Call minmax
        
     
    Next
    
End Sub
Sub Ticker_Volume()

Dim i As Long
Dim j As Integer
Dim Ticker As String
Dim Volume As Double
Dim lastrow As Long
Dim open_price As Double
Dim close_price As Double
Dim pricechange As Double
Dim percentchange As Double



lastrow = Cells(Rows.Count, "A").End(xlUp).Row
Volume = 0
open_price = 0
close_price = 0
pricechange = close_price - open_price

j = 2

For i = 2 To lastrow
     ' Are we in a new stock
     
     If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
       'Find open price
       
       
       open_price = Cells(i, 3).Value
       

     ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Ticker
        
        
        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        
        'Price change calculation
        
        
        close_price = Cells(i, 6).Value
        pricechange = close_price - open_price
        
         'In case a new stock is found
           
              If open_price = 0 Then
                percentchange = New_Ticker
              Else
              percentchange = pricechange / open_price
              End If
              
      
        
        
              Cells(j, 16).Value = pricechange
               Cells(j, 17).Value = percentchange
               Cells(j, 17).Style = "Percent"
              
        'Format percentages
        
               If Cells(j, 17).Value >= 0 Then
                  Cells(j, 17).Interior.ColorIndex = 4
               Else
                   Cells(j, 17).Interior.ColorIndex = 3
              End If
         'Store Values
        Cells(j, 15).Value = Ticker
        Cells(j, 18).Value = Volume
        
      ' Reset volume
      Volume = 0
      
      ' Increase row for storage
      
      j = j + 1

    
   Else
     Volume = Volume + Cells(i, 7).Value
  End If
  
       

 

Next i


End Sub


Sub minmax()
Dim Max_Inc As Double

Dim Max_Vol As Double




Dim Min_Inc As Double


Dim lastticker As Long
Dim MaxFind As Range
Dim MinFind As Range
Dim VolFind As Range




lastticker = Cells(Rows.Count, "O").End(xlUp).Row
Max_Inc = WorksheetFunction.Max(Range("q2", Range("q" & lastticker)))
Min_Inc = WorksheetFunction.Min(Range("q2", Range("q" & lastticker)))
Max_Vol = WorksheetFunction.Max(Range("R2", Range("R" & lastticker)))

Cells(2, 23).Value = Max_Inc
Cells(2, 23).Style = " Percent"

Cells(3, 23).Value = Min_Inc
Cells(3, 23).Style = " Percent"

Cells(4, 23).Value = Max_Vol



For i = 2 To lastticker
     If Cells(i, 17) = Max_Inc Then
        Cells(2, 22).Value = Cells(i, 15)
    ElseIf Cells(i, 17) = Min_Inc Then
       Cells(3, 22).Value = Cells(i, 15)
   ElseIf Cells(i, 18) = Max_Vol Then
       Cells(4, 22).Value = Cells(i, 15)

    Else
 
 End If
 
 Next i
 
 End Sub
 

Sub Clear()

'
' Macro7 Macro
'

'
    Columns("O:X").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("2015").Select
    Columns("O:AE").Select
    Range("O243").Activate
    Selection.Delete Shift:=xlToLeft
    Sheets("2014").Select
    Columns("O:AF").Select
    Range("O263").Activate
    Selection.Delete Shift:=xlToLeft
    
End Sub


