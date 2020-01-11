Attribute VB_Name = "Module1"
Sub stockAnalysis()

'Describing variables' data types

Dim dateOpen As Variant
Dim dateClose As Variant
Dim i As Double
Dim x As Double

'Creating Headers

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Creating variables

x = 2
i = 2

Cells(x, 9).Value = Cells(x, 1).Value

dateOpen = Cells(i, 3).Value

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Iteration to calculate total volume and keep close value

For i = 2 To lastRow

    If Cells(i, 1).Value = Cells(x, 9).Value Then

        volume = volume + Cells(i, 7).Value

        dateClose = Cells(i, 6).Value

    Else

        Cells(x, 10).Value = dateClose - dateOpen

    If dateClose <= 0 Then
        
        Cells(x, 11).Value = 0
                
    Else
        
        Cells(x, 11).Value = (dateClose / dateOpen) - 1

End If


'Formatting percent change column

    Cells(x, 11).Style = "Percent"
                    
    If Cells(x, 10).Value >= 0 Then
                            
        Cells(x, 10).Interior.ColorIndex = 4
                                
    Else
                            
        Cells(x, 10).Interior.ColorIndex = 3
                
    End If

Cells(x, 12).Value = volume

'Reset variables

dateOpen = Cells(i, 3).Value

volume = Cells(i, 7).Value

x = x + 1

Cells(x, 9).Value = Cells(i, 1).Value

End If

Next i

End Sub


