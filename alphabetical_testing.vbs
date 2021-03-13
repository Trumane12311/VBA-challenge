VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockmarket():

    Dim totalvolume As Double
    Dim RowCount As Long
    Dim RCount As Integer
    Dim startvalue As Double
    Dim closevalue As Long
    Dim change As Double
    Dim percentchange As Double
    Dim ws As Worksheet

    For Each ws In Worksheets
    'Run For Loop here for all sheets
    
        ws.Cells(1, 10).Value = "Ticker"
        
        ws.Cells(1, 13).Value = "Total Volume"
        
        ws.Cells(1, 11).Value = "Yearly Change"
        
        ws.Cells(1, 12).Value = "Percent Change"
        
        
        totalvolume = 0
        RCount = 2
        startvalue = 2
        
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        For i = 2 To RowCount
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                totalvolume = totalvolume + ws.Cells(i, 7).Value
                If ws.Cells(startvalue, 3) = 0 Then
                    For firstnonzero = startvalue To i
                        If ws.Cells(firstnonzero, 3).Value <> 0 Then
                            startvalue = firstnonzero
                            Exit For
                        End If
                        Next firstnonzero
                End If
                      change = ws.Cells(i, 6) - ws.Cells(startvalue, 3)
                      percentchange = change / ws.Cells(startvalue, 3)
                ws.Range("J" & RCount).Value = ws.Cells(i, 1).Value
                ws.Range("M" & RCount).Value = totalvolume
                ws.Range("K" & RCount).Value = change
                ws.Range("L" & RCount).Value = percentchange
                    
                totalvolume = 0
                RCount = RCount + 1
                change = 0
                
            Else
                totalvolume = totalvolume + ws.Cells(i, 7).Value
                End If
            
            Next i
            
                RowCount = ws.Cells(Rows.Count, "K").End(xlUp).Row
                For i = 1 To RowCount
                
                    If ws.Cells(i + 1, 11).Value >= 0 Then
                        ws.Cells(i + 1, 11).Interior.ColorIndex = 4
                    ElseIf ws.Cells(i + 1, 11).Value < 0 Then
                        ws.Cells(i + 1, 11).Interior.ColorIndex = 3
                    Else: ws.Cells(i + 1, 11).Interior.ColorIndex = 2
                End If
                
            Next i
                
                For i = 1 To RowCount
                ws.Cells(i + 1, 12).NumberFormat = "0.00%"
        
            Next i

    Next ws

End Sub


