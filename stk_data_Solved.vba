Attribute VB_Name = "Module1"
Sub stockTotals_YC_PC() 'works :-)
Dim book As Workbook
Dim sheet As Worksheet
Dim EOYCloserowNum As Double
Dim BOYOpen As Double
Dim EOYClose As Double
Dim yrlyChng As Double
Dim prcntChnge As Variant
Dim prcntChng As Long
Dim i As Double
    Dim c As Range
    Dim d As Range
    Dim maxStkvol As Double
    Dim grtPrcntInc As Double
    Dim grtPrcntDec As Double
    
  
For Each book In Workbooks
  For Each sheet In book.Worksheets
       MsgBox (sheet.Name)
       'MsgBox ("Rowcount is:" & sheet.Cells(Rows.Count, 1).End(xlUp).Row)
   EOYCloserowNum = 1
      For i = 2 To sheet.Cells(Rows.Count, 1).End(xlUp).Row
      'For i = 2 To 21
        'MsgBox ("Now checking the row :" & i & " with " & i - 1)
             
             If sheet.Cells(i, 1).Value <> sheet.Cells(i - 1, 1).Value Then
                totalVolrow = Cells(Rows.Count, 12).End(xlUp).Row + 1
                Cells(totalVolrow, 9) = sheet.Cells(i, 1).Value ' ticker name -> I
                Cells(totalVolrow, 12) = sheet.Cells(i, 12).Value 'calculating & updating totalstock vol -> L
                'MsgBox ("opn price of ticker: " & sheet.Cells(i, 1).Value & "is" & sheet.Cells(i, 3).Value)
                Cells(totalVolrow, 13) = sheet.Cells(i, 3).Value 'BOYOpen ->M
                BOYOpen = sheet.Cells(i, 3).Value 'collecting annualOpen value into a variable
                'MsgBox ("BOYOpen of " & sheet.Cells(i, 1).Value & " is " & BOYOpen)
                EOYCloserowNum = EOYCloserowNum + 1
                'MsgBox ("eoycloserownum: " & EOYCloserowNum) 'counting rownumber value. debugger step
                'MsgBox ("annualopen " & BOYOpen)
                'MsgBox ("annualclose " & EOYClose)
              Else
                 EOYCloserowNum = EOYCloserowNum + 1
                 Cells(totalVolrow, 12) = Cells(totalVolrow, 12) + sheet.Cells(i, 7).Value 'calculating & updating totalstock vol -> L
                 'MsgBox ("eoycloserownum: " & EOYCloserowNum) 'counting rownumber value. debugger step
                 Cells(totalVolrow, 14).Value = sheet.Cells(EOYCloserowNum, 6).Value 'EOYClose ->N
                 EOYClose = sheet.Cells(EOYCloserowNum, 6).Value 'collecting annualclose value into a variable
             
                 yrlyChng = EOYClose - BOYOpen
                 Cells(totalVolrow, 10).Value = yrlyChng
                   
                   If BOYOpen <> 0 Then
                    prcntChng = (yrlyChng / BOYOpen) * 100
                   Else
                    'MsgBox ("Denominator is 0. Cannot perform div operation to find percentchange")
                    prcntChnge = CVErr(xlErrDiv0)
                    Cells(totalVolrow, 11).Value = prcntChnge
                   End If
                    Cells(totalVolrow, 11).Value = prcntChng
                      
                      
                    If prcntChng < 0 Then
                     Cells(totalVolrow, 11).Interior.ColorIndex = 3
                    Else
                     Cells(totalVolrow, 11).Interior.ColorIndex = 4
                    End If
                 
               End If
     Next i
      If i Mod 5000 = 0 Then
      'MsgBox ("5000 rows completed")
      End If
      
  Next sheet
  'MsgBox ("All done with total stock volumes, yrly change and percent change")
    
    Set c = Range("L2:L" & Rows.Count)
    Set d = Range("K2:K" & Rows.Count)
    'MsgBox Application.WorksheetFunction.max(c)
    'MsgBox Application.WorksheetFunction.max(d)
    'MsgBox Application.WorksheetFunction.Min(d)
   
   grtPrcntInc = Range("L1").Application.WorksheetFunction.Max(d)
   Cells(2, 17).Value = grtPrcntInc
   grtPrcntDec = Range("L1").Application.WorksheetFunction.Min(d)
   Cells(3, 17).Value = grtPrcntDec
   maxStkvol = Range("L1").Application.WorksheetFunction.Max(c)
   Cells(4, 17).Value = maxStkvol
Next book

End Sub

