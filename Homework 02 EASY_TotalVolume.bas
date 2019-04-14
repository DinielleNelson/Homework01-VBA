Attribute VB_Name = "Module2"
Option Explicit

Sub EASY_TotalVolume()
    Dim i As Double
    Dim List_Row As Double
    Dim Abbrev As String
    Dim Vol_Total As Double
    Dim ws As Worksheet
    
  For Each ws In Worksheets
    ws.Activate
        
    Vol_Total = 0
    Range("J1") = "Ticker Abbrev"
    Range("K1") = "Volume Total"
    
    List_Row = 2
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            Abbrev = Cells(i, 1).Value
            
            Range("J" & List_Row).Value = Abbrev
            Range("K" & List_Row).Value = Vol_Total
        
            List_Row = List_Row + 1
            Vol_Total = 0
        Else
            Vol_Total = Vol_Total + Cells(i, 7).Value
        End If
    Next i
    
  Next ws
    
End Sub
