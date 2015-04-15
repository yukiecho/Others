Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("Driver Log Analysis.xlsm").Activate
Sheets("Input").Select
Sheets("Input").Range("A1").Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
     :=False, Transpose:=False

If Not IsOpen Then Wkb.Close False

'end here, refresh all contents in 'input' sheet

Wkb.Sheets("DailySummary").Range("A3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete

'step 1 delete all current contents in DailySummary

'Dim xr As Range
'Dim i As Integer
'Dim xdate As Date
'Dim startdate As Date
'Dim Enddate As Date
    
'change the date
startdate = #9/2/2014#
Enddate = #10/15/2014#

For xdate = startdate To Enddate
    Sheets("Reference").Range("A6:A27").Copy
    Sheets("DailySummary").Select
    ActiveCell.Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
        For i = 1 To 22
            ActiveCell.Value = xdate
            ActiveCell.Offset(1, 0).Range("A1").Select
        Next i

Next xdate

'step 2 figure out the 1st & last date in the records

'step 3 copy all the dates N times based on drivers names

'delete all the #NA lines






MsgBox "Input sheet and DailySummary have been updated successfully.", vbInformation

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub





End Sub
Function trackMinimum(rowRange As Range) As Date
    On Error Resume Next
    Dim j As Integer, minValue As Date
    Dim t0 As Double, t1 As Double
    Dim ans As Date

    t0 = CDbl(DateSerial(2000, 1, 1))
    t1 = CDbl(DateSerial(2100, 12, 31))
    ans = 0
    For j = 1 To rowRange.Columns.Count
        If ans = 0 Then ' You need to store the first valid value
            If rowRange.Cells(1, j).Value >= t0 And rowRange.Cells(1, j) <= t1 Then
                ans = rowRange.Cells(1, j).Value
            End If
        Else
            If (rowRange.Cells(1, j).Value >= t0 And rowRange.Cells(1, j) <= t1) _
               And rowRange.Cells.Value < ans Then
                ans = rowRange.Cells(1, j).Value
            End If
        End If
    Next j
    trackMinimum = ans
End Function
