Option Explicit
Sub StockTest()

Dim i As Long
Dim Ticker As String
Dim oValue As Double
Dim hValue As Double
Dim lValue As Double
Dim cValue As Double
Dim tVol As Double
Dim chValue As Double
Dim chValueCur As Double
Dim chValuePer As Double
Dim chValueP As Double
Dim volStart As Double
Dim volEnd As Double
Dim volTotal As Double
Dim volRange As Range
Dim VolStartAddr As String
Dim VolEndAddr As String
Dim Summary_Table_Ticker_Row As Integer
Dim Summary_Table_Change_Row As Integer
Dim Summary_Table_ChangeVol_Row As Integer
Dim Summary_Table_Vol_Row As Integer
Dim lastRow As Double
Dim StockCaptured As Boolean

lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range("H1").Value = "Ticker"
Range("I1").Value = "Yearly Change Value"
Range("J1").Value = "Percent change"
Range("K1").Value = "Total Volume"
oValue = 0
hValue = 0
lValue = 0
cValue = 0
tVol = 0
Summary_Table_Ticker_Row = 2
Summary_Table_Change_Row = 2
Summary_Table_ChangeVol_Row = 2
Summary_Table_Vol_Row = 2

    For i = 2 To lastRow
    If StockCaptured = False Then
        oValue = Cells(i, 3).Value
        volStart = Cells(i, 7).Value
        VolStartAddr = Cells(i, 7).Address
        StockCaptured = True
    End If
    
    
    
        'check if we are still on the same ticker
        'oValue = Cells(i, 3).Value
        'Range("I" & Summary_Table_Ticker_Row).Value = oValue
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
        cValue = Cells(i, 6).Value
        volEnd = Cells(i, 7).Value
        VolEndAddr = Cells(i, 7).Address
        chValue = cValue - oValue
        'MsgBox (VolStartAddr)
        'MsgBox (VolEndAddr)
        Set volRange = Range(VolStartAddr & ":" & VolEndAddr)
        volTotal = Application.WorksheetFunction.Sum(volRange)
        chValueP = chValue / oValue
        'Show Ticker value in column H
        Range("H" & Summary_Table_Ticker_Row).Value = Ticker
        'Show Change value in column I
        Range("I" & Summary_Table_Ticker_Row).Value = FormatCurrency(chValue, 2)
        'Show Percent Change value in column J
        Range("J" & Summary_Table_Ticker_Row).Value = FormatPercent(chValueP, 2, vbTrue)
        Range("J" & Summary_Table_Ticker_Row).NumberFormat = "0.00%;[Red]-0.00%"
        'Show total volume in column K
        Range("K" & Summary_Table_Ticker_Row).Value = volTotal
        'Add 1 to the Summary Row
        Summary_Table_Ticker_Row = Summary_Table_Ticker_Row + 1
        StockCaptured = False
        End If
    Next i
    
Columns("A:K").EntireColumn.AutoFit
    
End Sub
