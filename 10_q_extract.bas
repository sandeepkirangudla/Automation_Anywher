Attribute VB_Name = "Module1"
Sub form1()
Attribute form1.VB_Description = "form\n"
Attribute form1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' form1 Macro
' form
'

'
    ChDir "C:\Users\gsand\OneDrive\Desktop\Automation Anywhere\Project_dev\forms"
    Workbooks.OpenText Filename:= _
        "C:\Users\gsand\OneDrive\Desktop\Automation Anywhere\Project_dev\forms\form_2015_1.txt" _
        , Origin:=437, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(9, 1), Array(74, 1), Array(83, 1), Array(96, 1)), _
        TrailingMinusNumbers:=True
    Rows("1:8").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$318648").AutoFilter Field:=1, Criteria1:="=10-Q" _
        , Operator:=xlOr, Criteria2:="=10-Q/A"
    Range("A1:E318648").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Columns("E:E").EntireColumn.AutoFit
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "10-q"
    Sheets("form_2015_1").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Range("E14").Select
    ActiveWorkbook.Save
End Sub
