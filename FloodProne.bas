Attribute VB_Name = "FloodProne"
Sub FPnewest2()
Attribute FPnewest2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FPnewest2 Macro
'

'
    Range("D3:E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Background").Select
    Range("B2").Select
    ActiveSheet.Paste
    Columns("F:G").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$F$1:$G$1048575").RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlYes
    ActiveWindow.LargeScroll Down:=6
    Range("F301:G301").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-27
    ActiveWindow.ScrollRow = 1
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H2:H300").Select
    ActiveWindow.ScrollRow = 1
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Background").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Background").Sort.SortFields.Add Key:=Range("H2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Background").Sort
        .SetRange Range("H2:H300")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("K:L").Select
    ActiveSheet.Range("$K$1:$L$1048575").RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlYes
    Range("Y2").Select
    ActiveSheet.ChartObjects("Chart 5").Activate
    ActiveSheet.ChartObjects("Chart 5").Activate
    Range("R2:R300").Select
    Selection.Copy
    Range("S1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("S1:S299").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Background").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Background").Sort.SortFields.Add Key:=Range("S1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Background").Sort
        .SetRange Range("S1:S299")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("X5").Select
    Selection.Copy
    Sheets("Input-Results").Select
    Range("A18").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C30").Select
    Sheets("Background").Select
    Range("U1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Input-Results").Select
    Range("A10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Background").Select
    Range("U3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Input-Results").Select
    Range("A12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Background").Select
    Range("U5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Input-Results").Select
    Range("A13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Background").Select
    Range("U7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Input-Results").Select
    Range("A14").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Background").Select
    Range("U9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Input-Results").Select
    Range("A15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K1").Select
    Application.CutCopyMode = False
End Sub
