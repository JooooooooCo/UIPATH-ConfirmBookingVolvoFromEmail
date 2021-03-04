Attribute VB_Name = "ToLogMove"
Sub MoveToLog()
Attribute MoveToLog.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MOVE ROW 3 FROM RUN TO LOG SHEET
'

'

    Sheets("Run").Select
    Rows("3:3").Select
    Selection.Copy
    Sheets("Log").Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Sheets("Run").Select
    Application.CutCopyMode = False
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Range("A3").Select


End Sub
