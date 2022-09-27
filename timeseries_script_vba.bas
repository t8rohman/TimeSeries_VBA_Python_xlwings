Attribute VB_Name = "Module1"
Sub SampleCall()
' only a sample template from xlwings quickstart, used for quick debugging
    mymodule = Left(ThisWorkbook.name, (InStrRev(ThisWorkbook.name, ".", -1, vbTextCompare) - 1))
    RunPython "import " & mymodule & ";" & mymodule & ".main()"
End Sub

Sub data_diagnose()
' calling the data_diagnose() function in the module
    mymodule = Left(ThisWorkbook.name, (InStrRev(ThisWorkbook.name, ".", -1, vbTextCompare) - 1))
    RunPython "import " & mymodule & ";" & mymodule & ".data_diagnose()"
    Sheets("diagnose").Select
End Sub

Sub data_forecast_hwes3()
    mymodule = Left(ThisWorkbook.name, (InStrRev(ThisWorkbook.name, ".", -1, vbTextCompare) - 1))
    RunPython "import " & mymodule & ";" & mymodule & ".forecast_hwes3()"
    Sheets("forecast").Select
    Cells.Select
    With Selection.Font  ' used for standard formatting, Arial 10
        .name = "Arial"
        .Size = 10
    End With
    Range("A1").Select
End Sub

Sub data_forecast_hwes2()
    mymodule = Left(ThisWorkbook.name, (InStrRev(ThisWorkbook.name, ".", -1, vbTextCompare) - 1))
    RunPython "import " & mymodule & ";" & mymodule & ".forecast_hwes2()"
    Sheets("forecast").Select
    Cells.Select
    With Selection.Font  ' used for standard formatting, Arial 10
        .name = "Arial"
        .Size = 10
    End With
    Range("A1").Select
End Sub

Sub data_forecast_hwes1()
    mymodule = Left(ThisWorkbook.name, (InStrRev(ThisWorkbook.name, ".", -1, vbTextCompare) - 1))
    RunPython "import " & mymodule & ";" & mymodule & ".forecast_hwes1()"
    Sheets("forecast").Select
    Cells.Select
    With Selection.Font  ' used for standard formatting, Arial 10
        .name = "Arial"
        .Size = 10
    End With
    Range("A1").Select
End Sub

Sub main_openfile()
' used to open and paste the data from selected file
Dim UserRange As Range, DefaultRange As String, filename As Variant, nr As Integer, min_obs As Double
Dim tWB As Workbook, aWB As Workbook
On Error Resume Next

Set tWB = ThisWorkbook

filename = Application.GetOpenFilename(Title:="Select Your File")
' Error handling
If filename = False Then
    Exit Sub
End If

MsgBox ("Select the range you want to import, consisting of time period and the observed level. Be sure to put 2 rows, with period on left and observed.")

Workbooks.Open filename
DefaultRange = Selection.Address ' Selection before subroutine is executed

Set UserRange = Application.InputBox(Prompt:="Select a range to copy to the main sheet!", Title:="Instruction", default:=DefaultRange, Type:=8)
Err.Clear
' Error handling
If UserRange Is Nothing Then
    Set aWB = ActiveWorkbook
    aWB.Close SaveChanges:=False
    Exit Sub
ElseIf UserRange.Columns.Count <> 2 Then
    MsgBox ("The data should be in 2 columns. Press the button again to retry!")
    Set aWB = ActiveWorkbook
    aWB.Close SaveChanges:=False
    Exit Sub
End If

Set aWB = ActiveWorkbook
aWB.Worksheets(1).Activate
UserRange.Select
Selection.Copy
aWB.Close SaveChanges:=False

tWB.Worksheets("data").Activate
Range("B2").Select
ActiveSheet.Paste
Range("A1").Select

nr = Cells(Rows.Count, 2).End(xlUp).Row
Range("A2:A4").AutoFill Destination:=Range("A2:A" & nr), Type:=xlFillSeries

' modifying the data in chart

min_obs = WorksheetFunction.Min(Range("C:C"))
ActiveSheet.ChartObjects("main_chart").Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MinimumScale = min_obs - min_obs / 10

Range("A1").Select

End Sub

Sub main_reset()
' reset all data from data sheet
Dim nr As Integer

nr = Cells(Rows.Count, 2).End(xlUp).Row
If WorksheetFunction.CountA(Range("B2:C" & nr)) = 0 Then
    Exit Sub
End If
ActiveSheet.ListObjects("Table1").Resize Range("$A$1:$C$4")
Range("A5:C" & nr).Clear
Range("B2:C4").Clear

End Sub

Sub reset_diagnose()
Attribute reset_diagnose.VB_ProcData.VB_Invoke_Func = " \n14"
' reset all data from diagnose sheet
    On Error GoTo ErrorHandler
    Sheets("diagnose").Select
    ActiveSheet.Shapes.Range(Array("decompose_chart", "hwes1_chart", _
        "hwes2_chart", "hwes3_chart")).Select
    Selection.Delete
    Exit Sub

ErrorHandler:
    MsgBox "All data is cleared."
    Exit Sub
End Sub

Sub reset_forecast()
' reset all data from forecast sheet
    On Error GoTo ErrorHandler
    Sheets("forecast").Select
    Range("A2:F10000").Clear
    Range("J1:M3").Clear
    Call cleaning_eval_metric_forecast
    ActiveSheet.Shapes.Range(Array("chart_forecast")).Select
    Selection.Delete
    Exit Sub

ErrorHandler:
    MsgBox "All data is cleared."
    Exit Sub
End Sub

Sub cleaning_eval_metric_forecast()
Attribute cleaning_eval_metric_forecast.VB_ProcData.VB_Invoke_Func = " \n14"
' only for cleaning the evaluation metrics in forecast sheet
' using the call function in vba
    Range("J1:M1").Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    Selection.Merge
    Range("J2:M2").Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    Selection.Merge
    Range("J3:M3").Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    Selection.Merge
    Range("H1:M3").Select
    Range("J1").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
End Sub
