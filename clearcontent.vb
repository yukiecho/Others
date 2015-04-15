Sub kTest()
    Dim ws As Worksheet
     
    For Each ws In Worksheets
    'Select Case ws.Name
    'Case "September Summary", "Driver's Rate"

    'Case Else
        Application.ScreenUpdating = False
        ws.Range("B2:C8,B10:C16,B18:C24,B26:C32,B34:C40,F2:G8,F10:G16,F18:G24,F26:G32,F34:G40").ClearContents

    'End Select
    Application.ScreenUpdating = True
Next ws

    
End Sub
