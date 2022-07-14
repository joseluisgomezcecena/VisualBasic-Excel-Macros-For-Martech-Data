Dim ws As Worksheet

Application.DisplayAlerts = False
For Each ws In ThisWorkbook.Worksheets
    If ws.CodeName <> "Sheet1" Then ws.Delete
Next
Application.DisplayAlerts = True