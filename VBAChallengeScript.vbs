Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim k As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For k = 1 To WS_Count
         

    ' Declare Variables

    Dim si As Integer

    'si calculates output row

    Dim opening As Double
    Dim closing As Double
    Dim change As Double
    Dim total As LongLong

    'Assign titles of columns

    Worksheets(k).Range("J1").Value = "Ticker Name"
    Worksheets(k).Range("K1").Value = "Yearly Change"
    Worksheets(k).Range("L1").Value = "Percent Change"
    Worksheets(k).Range("M1").Value = "Total Stock Volume"
    Worksheets(k).Range("Q1").Value = "Ticker"
    Worksheets(k).Range("R1").Value = "Value"
    Worksheets(k).Range("P2").Value = "Greatest % Increase"
    Worksheets(k).Range("P3").Value = "Greatest % Decrease"
    Worksheets(k).Range("P4").Value = "Greatest Total Volume"

    total = 0
    
    'Start si as 2
        si = 2
        
    ' Define last row
    lastrow = Worksheets(k).Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Define initial opening
    
        opening = Worksheets(k).Cells(2, 3).Value
        
        For I = 2 To lastrow
        
             'Total Value of stock
             ' total = total + Worksheets(k).Cells(si, "G").Value
            total = total + Worksheets(k).Cells(I, "G").Value
        
              ' To assign the ticker to row J and check if next value is same as previous
        
             If Worksheets(k).Cells(I, 1).Value <> Worksheets(k).Cells(I + 1, 1).Value Then
                Worksheets(k).Cells(si, "J").Value = Worksheets(k).Cells(I, 1).Value
                
                'Assign Total Volume to Cells
                
                Worksheets(k).Cells(si, "M").Value = total
                
                'Set closing variable
                
                closing = Worksheets(k).Cells(I, 6).Value
                
                'Difference between opening and closing
                
                Worksheets(k).Cells(si, "K").Value = closing - opening
                
                
                'percent change is larger - smaller/ first
      
                change = (closing - opening) / opening
      
                Worksheets(k).Cells(si, "L").Value = change
                
                If Worksheets(k).Cells(si, "K") > 0 Then Worksheets(k).Cells(si, "K").Interior.ColorIndex = 4
                If Worksheets(k).Cells(si, "K") < 0 Then Worksheets(k).Cells(si, "K").Interior.ColorIndex = 3
                       
        
                'Reassign opening for next loop
                
                opening = Worksheets(k).Cells(I + 1, 3).Value
                
            
              si = si + 1
              total = 0
              
            End If
        
        Next I
    
    
    
    'New loop for greatest values
    
    'Declare new variables
    
    Dim greatestincreasename As String
    Dim greatestdecreasename As String
    Dim greatestvolumename As String
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As LongLong
    
    'Define new last row
    lastrow_2 = Worksheets(k).Cells(Rows.Count, "L").End(xlUp).Row
    
    
    
    'Start new loop
    For I = 2 To lastrow_2
    
    
        If Worksheets(k).Cells(I, "L").Value > greatestincrease Then
    
            greatestincrease = Worksheets(k).Cells(I, "L").Value
            greatestincreasename = Worksheets(k).Cells(I, "J").Value
    
    
        End If
    
        If Worksheets(k).Cells(I, "L").Value < greatestdecrease Then
    
            greatestdecrease = Worksheets(k).Cells(I, "L").Value
            greatestdecreasename = Worksheets(k).Cells(I, "J").Value
            
            
        End If
        
        If Worksheets(k).Cells(I, "M").Value > greatestvolume Then
    
            greatestvolume = Worksheets(k).Cells(I, "M").Value
            greatestvolumename = Worksheets(k).Cells(I, "J").Value
            
        End If
        
    Next I
    
    'Assign locations for increase and decrease variables
    
    Worksheets(k).Range("Q2").Value = greatestincreasename
    Worksheets(k).Range("R2").Value = greatestincrease
    Worksheets(k).Range("Q3").Value = greatestdecreasename
    Worksheets(k).Range("R3").Value = greatestdecrease
    Worksheets(k).Range("Q4").Value = greatestvolumename
    Worksheets(k).Range("R4").Value = greatestvolume
        
        
    'Change column widths
    
    Worksheets(k).Cells.EntireColumn.AutoFit
    Worksheets(k).Range("L2:L" & lastrow).NumberFormat = "0.00%"
    Worksheets(k).Range("R2:R3").NumberFormat = "0.00%"
    
    
                ' The following line shows how to reference a sheet within
                ' the loop by displaying the worksheet name in a dialog box.
                'MsgBox Worksheets(k).Name

         Next k

      End Sub
