Attribute VB_Name = "Module1"
Option Explicit

Sub HARD_GreatIncrease()
    Dim i As Double
    Dim List_Row As Double
    Dim Vol_Total As Double
    Dim LastRow As Double
    Dim Abbrev As String
    Dim Opening As Double
    Dim Closing As Double
    Dim Difference As Double
    Dim Percent As Double
    Dim Perc_Min As Double
    Dim Perc_Max As Double
    Dim Vol_Max As Double
    Dim Perc_rng As Range
    Dim Vol_rng As Range
    Dim x As Integer
    Dim holder As String
    Dim ws As Worksheet
        
    
  For Each ws In Worksheets
    ws.Activate
        
    Vol_Total = 0
    Range("J1") = "Ticker Abbrev"
    Range("K1") = "Difference"
    Range("L1") = "Percent"
    Range("M1") = "Volume Total"
    
    List_Row = 2
    Opening = Cells(2, 3).Value
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            Abbrev = Cells(i, 1).Value
            Closing = Cells(i, 6).Value
            Difference = Closing - Opening
            
            If Opening = 0 Then
                Percent = ((Closing / 1) - 1) * 100
            ElseIf Difference <> 0 Then
                Percent = ((Closing / Opening) - 1) * 100
            Else: Percent = 0
            End If
            
            Range("J" & List_Row).Value = Abbrev
            Range("K" & List_Row).Value = Difference
            
            If Percent > 0 Then
                Range("K" & List_Row).Interior.ColorIndex = 10
            ElseIf Percent < 0 Then
                Range("K" & List_Row).Interior.ColorIndex = 30
            End If
            
            Range("L" & List_Row).Value = Percent
            Range("M" & List_Row).Value = Vol_Total
            
            Opening = Cells(i + 1, 3).Value
            List_Row = List_Row + 1
            Vol_Total = 0
        Else
            Vol_Total = Vol_Total + Cells(i, 7).Value
        End If
    Next i
    
    
    Set Perc_rng = Range("L:L")
    Set Vol_rng = Range("M:M")
    
    Perc_Max = Application.WorksheetFunction.Max(Perc_rng)
    Perc_Min = Application.WorksheetFunction.Min(Perc_rng)
    Vol_Max = Application.WorksheetFunction.Max(Vol_rng)


    Range("O1") = "Greatest % increase"
    Range("O2") = "Greatest % Decrease"
    Range("O3") = "Greatest total volume"
    
    Range("P1") = Perc_Max
    Range("P2") = Perc_Min
    Range("P3") = Vol_Max
    

    
        For x = 2 To Cells(Rows.Count, 12).End(xlUp).Row
            If Range("L" & x) = Range("P1").Value Then
                holder = Range("J" & x).Value
                Range("Q1") = holder
            End If
        Next x
        
        For x = 2 To Cells(Rows.Count, 12).End(xlUp).Row
            If Range("L" & x) = Range("P2").Value Then
                holder = Range("J" & x).Value
                Range("Q2") = holder
            End If
        Next x
        
        For x = 2 To Cells(Rows.Count, 12).End(xlUp).Row
            If Range("M" & x) = Range("P3").Value Then
                holder = Range("J" & x).Value
                Range("Q3") = holder
            End If
        Next x
            
    
  Next ws
    
End Sub
