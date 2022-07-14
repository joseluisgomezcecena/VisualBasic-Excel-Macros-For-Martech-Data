'Code Created by Sumit Bansal from trumpexcel.com
Sub SplitEachWorksheet()
Dim FPath As String
FPath = Application.ActiveWorkbook.Path
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each ws In ThisWorkbook.Sheets
    ws.Copy
    nospace = Trim(ws.Name)
    myfilename = Replace(nospace, vbLf, "")
    Application.ActiveWorkbook.SaveAs Filename:=FPath & "\" & myfilename & ".xlsx"
    Application.ActiveWorkbook.Close False
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
