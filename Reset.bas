Attribute VB_Name = "Reset"
Sub ResetNewest1()
Attribute ResetNewest1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ResetNewest1 Macro
'

'
    Range("A10").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A13").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A14").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A15").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A18").Select
    ActiveCell.FormulaR1C1 = ""
    Range("C34").Select
    Sheets("Background").Select
    Columns("S:S").Select
    Selection.ClearContents
    Range("B2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("F2:G2").Select
    Selection.AutoFill Destination:=Range("F2:G300"), Type:=xlFillDefault
    Range("F2:G300").Select
    Range("J299").Select
    ActiveWindow.SmallScroll Down:=-27
    ActiveWindow.ScrollRow = 1
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("K2:L2").Select
    Selection.AutoFill Destination:=Range("K2:L300"), Type:=xlFillDefault
    Range("K2:L300").Select
    Range("N308").Select
    ActiveWindow.ScrollRow = 1
    Sheets("Input-Results").Select
    Range("B3").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B4").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B5").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B6").Select
    ActiveCell.FormulaR1C1 = ""
    Range("D3:E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("H19").Select
End Sub
Sub FPnewest1()
Attribute FPnewest1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FPnewest1 Macro
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
    ActiveWindow.ScrollRow = 1
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Background").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Background").Sort.SortFields.Add Key:=Range("H2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Background").Sort
        .SetRange Range("H2:H21")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("K:L").Select
    ActiveSheet.Range("$K$1:$L$1048575").RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlYes
    Range("R2:R300").Select
    Selection.Copy
    Range("S1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
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
    ActiveSheet.ChartObjects("Chart 5").Activate
    ActiveSheet.ChartObjects("Chart 5").Activate
    Range("X5").Select
    Selection.Copy
    Sheets("Input-Results").Select
    Range("A18").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C32").Select
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
    Range("I26").Select
End Sub
