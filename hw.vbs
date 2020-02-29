

Sub WorksheetLoop()
             '-------------------------------------------------------------------------------
                      'Part I
             
             'Create a script that will loop through all the stocks for one year for each run and take the following information.
             '-> The ticker symbol.
             '-> Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
             '-> The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
             '-> The total stock volume of the stock."
                        
             
             '-------------------------------------------------------------------------------
             
             
    'if finds a div/0 skip that error an continue the iteration
    On Error Resume Next
    Application.ScreenUpdating = False

    ' variable for:'
    'the worksheets,counting the existing worksheets, to store the rows in the new table, to store the sum of the values that are equal
    Dim WS_Count, I, sumary_coun, volume As Double
    'variable that stores the tickers available
    Dim u_ticker As String
    'variable for the rows
    Dim j As Double
    'variables for the news columns
    Dim open_price As Double
    Dim close_price As Double
    'varibale for the yearly change
    Dim ychange As Double
    Dim pchange As Double
    'variable for range
    Dim rng As Range
    'variable ofr conditions
    Dim condition1 As FormatCondition, condition2 As FormatCondition
    Dim lRow, lCol, lRow2 As Double
    Dim f As Double
         
    'initializing variables
         
    volume = 0
    j = 2
         
    ' Set WS_Count equal to the number of worksheets in the active workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

 ' Begin tho go through all ws .
     For I = 1 To WS_Count
         ' activate the sheet
         Worksheets(I).Select
         sumary_count = 2
         ychange = 0
         pchange = 0
                 
            
         'Find the last non-blank cell in column A(1)
         lRow = Cells(Rows.Count, 1).End(xlUp).Row
         'MsgBox Str(lRow)
         'Find the last non-blank cell in row 1
         lCol = Cells(1, Columns.Count).End(xlToLeft).Column
         'MsgBox Str(lCol)            
            
         'variable for ychange
         f = 2
         open_price = Cells(f, 3).Value
         'MsgBox Str(open_price)
          
            
         'loop for compare the rows in the column ticker
         For j = 2 To lRow             
             If Cells(j, 1).Value <> Cells(j + 1, 1) Then              
                 ' sume the valume
                 volume = volume + Cells(j, 7).Value
                 'MsgBox Str(volume)
                        
                 ' keep the tickers
                 u_ticker = Cells(j, 1).Value
                 'ychange
                 open_price = Cells(f, 3).Value
                     if open_price =0 Then
                        open_price= 0.0000001

                     end if
                                                                 
                 'MsgBox Str(open_price)
                 close_price = Cells(j, 6).Value
                   'MsgBox Str(close_price)
                 ychange = close_price - open_price
                  
                 'MsgBox Str(close_price) & "-" & Str(open_price)
                    
                 'percentage change
                 pchange = ((close_price - open_price) / Abs(open_price))
                                       
                 'MsgBox (pchange)
                                    
                 'Print the tickers in the new table
                 Range("J" & sumary_count).Value = u_ticker
                    
                 'Print the volume in the new table
                 Range("M" & sumary_count).Value = volume
                 Range("M" & sumary_count).NumberFormat = "#,#00#"
                 'print ychancge
                 'Cells(j, 11).Value = ychange
                 Range("K" & sumary_count).Value = ychange
                                    
                 'print percentage change
                 Range("L" & sumary_count).Value = pchange
                 Range("L" & sumary_count).NumberFormat = "0.000%"
                    
                 'extract from -> https://www.wallstreetmojo.com/vba-conditional-formatting/
                 Set rng = Range("L" & sumary_count)
                 Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
                 Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
                    
                    With condition1
                         .Interior.Color = RGB(174, 240, 194)
                     End With

                     With condition2
                         .Interior.Color = RGB(240, 128, 128)
                     End With
                                       
                 
                 'increase the sumary by 1
                 sumary_count = sumary_count + 1
                        
                 'clean the variable volume
                    
                  volume = 0
                
                    
                 'change the f variable that count the position in open price
                 f = j + 1
             Else
                 ' sume the valume
                 volume = volume + Cells(j, 7).Value
                 open_price = Cells(j, 3).Value
                close_price = Cells(j, 6).Value
                 ychange = close_price - open_price
                    
                   
                 'MsgBox Str(ychange)
             End If
                          
         Next j
             
             
             '-------------------------------------------------------------------------------
             
                                                'challenge part
                                                
             '-------------------------------------------------------------------------------
             
         lRow2 = Cells(Rows.Count, 10).End(xlUp).Row
         'For n = 2 To lRow2
         '    If Cells(n, 11).Value = 0 Then              
          '       Cells(n, 12).Value = 0
          '   End If
         'Next n
             
         Cells(2, 18).Value = "Greatest % Increase"
         Cells(3, 18).Value = "Greatest % Decrease"
         Cells(4, 18).Value = "Greatest Total Volumen"
             
         Cells(1, 19).Value = "Ticker"
         Cells(1, 20).Value = "Value"
         lMax = WorksheetFunction.Max((Range("L2:L500000")))
         lMax1 = lMax * 100
         'MsgBox Str(lMax1)
       
         lMin = WorksheetFunction.Min((Range("L2:L50000")))
         lMin1 = lMin * 100
         'MsgBox Str(lMin)
            
         lMaxV = WorksheetFunction.Max(Range("M2:M500000"))
         'MsgBox Str(lMaxV)
            
             
         For m = 2 To lRow2                                    
             If Cells(m, 12).Value = lMax Then
             
                ticker = Cells(m, 10).Value
                Cells(2, 19).Value = ticker
                Range("T2").Value = lMax1 / 100
                Range("T2").NumberFormat = "0.000%"
                
                'MsgBox ticker
                    
             ElseIf Cells(m, 12).Value = lMin Then
                
                 ticker = Cells(m, 10).Value
                 Cells(3, 19).Value = ticker
                 Range("T3").Value = lMin1 / 100
                 Range("T3").NumberFormat = "0.000%"
                    
             ElseIf Cells(m, 13).Value = lMaxV Then
                
                 ticker = Cells(m, 10).Value
                 Cells(4, 19).Value = ticker
                 Range("T4").Value = lMaxV
                 Range("T4").NumberFormat = "#,#00#"
                                      
                
             End If
                         
         Next m
                                                           
            'MsgBox ActiveWorkbook.Worksheets(i).Name
          
     Next I 'end for ws

End Sub


Sub clean_info()

         Dim WS_Count As Integer
         Dim I As Integer
         Dim rng As Range

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For I = 1 To WS_Count
            ' activate the sheet
            Worksheets(I).Select
            ' Insert your code here.
            
            Set rng = Range("J2:T500000")
            rng.Clear
            
            'MsgBox ActiveWorkbook.Worksheets(I).Name

        Next I
  End Sub


