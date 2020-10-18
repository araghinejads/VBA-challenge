Attribute VB_Name = "Module11"
Sub VBAofWallStrest()

   
'
' Repeat calculation for all Worksheets
'
   
   For Each ws In Worksheets

   
        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        ' Not necessary. just in case
        
        WorksheetName = ws.Name

  '
  ' Define headers for the first riw from coulumn 10 to 13
  '
  
   ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
     ws.Cells(1, 12).Value = "Percent Change"
      ws.Cells(1, 13).Value = "Total Stock Volume"
      
      
      '
      ' Define variables to start calculations
      '
      '
       
      counter = 1
      
      ' fisrst value of the "opening" variable
      
      
      Openingprice = ws.Cells(2, 3).Value
      
      ' an additional row is added at the end for ease of calculations
      
      ws.Cells(lastrow + 1, 2) = 0


'
'    Start the main Loops (form row 3 to the 1 row after the final row)
'



For i = 3 To lastrow + 1



' to determine the change of ticker variables, the data is chacked, whenever the date is less than the previous data it is recignized as the start of
' a new tikcer



If ws.Cells(i, 2).Value < ws.Cells(i - 1, 2).Value Then


        counter = counter + 1

        ws.Cells(counter, 10).Value = ws.Cells(i - 1, 1).Value
        ws.Cells(counter, 11).Value = difference

                ' Color code the percentages
                
                
                If ws.Cells(counter, 11).Value < 0 Then

                    ws.Cells(counter, 11).Interior.ColorIndex = 3

                Else

                ws.Cells(counter, 11).Interior.ColorIndex = 4

                End If


        ws.Cells(counter, 12).Value = DifferencePercent

        ws.Cells(counter, 12).NumberFormat = "0.00%"



        ws.Cells(counter, 13).Value = totalvol
        totalvol = 0
        Openingprice = ws.Cells(i, 3).Value

Else



totalvol = totalvol + ws.Cells(i, 7).Value
difference = ws.Cells(i, 6).Value - Openingprice

    
       ' To avoid divide by zero
    
    If Openingprice = 0 Then

            DifferencePercent = 0

            Else

            DifferencePercent = difference / Openingprice


      End If


End If


Next i




'
' Find the Maximim and minimum values
'


MaxChange = ws.Cells(2, 12).Value
MinChange = ws.Cells(2, 12).Value
Maxtotaltotalvol = ws.Cells(2, 13).Value

MAXTC = ws.Cells(2, 10).Value
MinTC = ws.Cells(2, 10).Value
MAXTV = ws.Cells(2, 10).Value


For k = 2 To lastrow



If ws.Cells(k, 12).Value > MaxChange Then

MaxChange = ws.Cells(k, 12).Value
MAXTC = ws.Cells(k, 10).Value

End If


If ws.Cells(k, 12).Value < MinChange Then

MinChange = ws.Cells(k, 12).Value
MinTC = ws.Cells(k, 10).Value

End If



If ws.Cells(k, 13).Value > Maxtotaltotalvol Then

Maxtotaltotalvol = ws.Cells(k, 13).Value
MAXTV = ws.Cells(k, 10).Value

End If



Next k


ws.Cells(1, 17).Value = "Values"

ws.Cells(2, 17).Value = MaxChange
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).Value = MinChange
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 17).Value = Maxtotaltotalvol


ws.Cells(1, 16).Value = "Ticker"
ws.Cells(2, 16).Value = MAXTC
ws.Cells(3, 16).Value = MinTC
ws.Cells(4, 16).Value = MAXTV



ws.Cells(2, 15).Value = "Max%Increase"
ws.Cells(3, 15) = "Max%Decrease"
ws.Cells(4, 15) = "MaxTotalValue"





Next ws
  






End Sub

